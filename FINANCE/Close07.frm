VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Close07 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recost Purchase Orders"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Selection"
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
      Left            =   8400
      TabIndex        =   20
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   2880
      Width           =   1920
   End
   Begin VB.CommandButton cmdSelPO 
      Caption         =   "Select PO's"
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
      Left            =   5880
      TabIndex        =   19
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   1320
      Width           =   1800
   End
   Begin VB.CheckBox chkRecostShipments 
      Caption         =   "Recost shipped PO items"
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CheckBox chkRecostPicks 
      Caption         =   "Recost purchased pick items"
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2040
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CheckBox chkDiagnose 
      Caption         =   "Diagnose only (do not update costs)"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2760
      Width           =   3735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7905
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Close07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   360
      Left            =   6480
      Picture         =   "Close07.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print The Report"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "Close07.frx":0938
      Height          =   350
      Left            =   6000
      Picture         =   "Close07.frx":0E12
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton cmdRecost 
      Caption         =   "Recost"
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
      Left            =   10560
      TabIndex        =   6
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   2880
      Width           =   1920
   End
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
      Left            =   5880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   1680
   End
   Begin VB.Frame fraDateRange 
      Height          =   1275
      Left            =   600
      TabIndex        =   13
      Top             =   660
      Width           =   5055
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "99999"
         Top             =   1080
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cmbCompletedThru 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Tag             =   "4"
         Top             =   660
         Width           =   1095
      End
      Begin VB.ComboBox cmbCompletedFrom 
         Height          =   315
         Left            =   3000
         TabIndex        =   0
         Tag             =   "4"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Recost a maximum of"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1110
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "Receipts"
         Height          =   255
         Left            =   3660
         TabIndex        =   16
         Top             =   1140
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Through"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recost items received from"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4215
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   3600
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7800
      Picture         =   "Close07.frx":12EC
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7800
      Picture         =   "Close07.frx":1676
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Recost Purchase Orders"
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
      Left            =   240
      TabIndex        =   11
      Top             =   300
      Width           =   5775
   End
End
Attribute VB_Name = "Close07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit

Dim bCantClose As Byte
Dim bOnLoad As Byte
Dim bGoodPrt As Byte
Dim bGoodRun As Byte
Dim bLotsOn As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCompletedFrom_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbCompletedThru_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 10
      If Grd.Row >= 1 Then
         If Grd.Row = 0 Then Grd.Row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
    End If
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      Grd.Col = 10
      
      If Grd.Row >= 1 Then
         If Grd.Row = 0 Then Grd.Row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
End Sub

Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 10
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next

End Sub

Private Sub cmdRecost_Click()
   
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   Dim iList As Long
   
   bCantClose = 0
   
   'get the list of all receipts between the requested dates
   Dim ia As New ClassInventoryActivity
   Dim sReturnMsg As String
   
   ia.LoggingEnabled = True
   'Dim rdo As ADODB.Recordset
   Dim success As Boolean

   ia.DiagnoseOnly = CBool(chkDiagnose.Value)
   ia.Log "Recosting purchase orders completed between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If ia.DiagnoseOnly Then
      ia.Log "Diagnosing Only.  POs will not be updated"
   End If
   ia.Log ""
   
   Dim totalItems As Integer
   Dim recostedItems As Integer
   
   Dim strPart As String
   Dim strPINum As String
   Dim strPIItem As String
   Dim strPIRev As String
   Dim strPIAQty As String
   Dim strPIAmt As String
   recostedItems = 1
   For iList = 1 To Grd.Rows - 1
      cmdRecost.enabled = False
      cmdCan.enabled = False
       
      
      Grd.Col = 10
      Grd.Row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
        
         Grd.Col = 0
         strPart = Trim(Grd.Text)
         Grd.Col = 1
         strPINum = Trim(Grd.Text)
         Grd.Col = 2
         strPIItem = Trim(Grd.Text)
         Grd.Col = 3
         strPIRev = Trim(Grd.Text)
         Grd.Col = 4
         strPIAQty = Trim(Grd.Text)
         Grd.Col = 5
         strPIAmt = Trim(Grd.Text)
         
         ia.PartNumber = strPart
         StatusBar1.SimpleText = " PO " & strPINum & " item " & _
                     strPIItem & strPIRev & " part " & strPart & " completed. - Recosted " & recostedItems

         If ia.UpdateReceiptCosts(strPart, CLng(strPINum), CInt(strPIItem), strPIRev, CDbl(strPIAQty), CDbl(strPIAmt), True) _
                  <> "UpdateReceiptCosts failed" Then
            recostedItems = recostedItems + 1
         End If
      End If
      Next
   
      Dim sMsg As String
      sMsg = "Recosted " & recostedItems & " PO's."
      ia.Log ""
      ia.Log sMsg
      StatusBar1.SimpleText = sMsg & "  See log"
      cmdRecost.enabled = True
      cmdCan.enabled = True
      MouseCursor ccArrow

End Sub

Private Sub old_cmdRecost_Click()
   
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   
   'get the list of all receipts between the requested dates
   Dim ia As New ClassInventoryActivity
   Dim sReturnMsg As String
'   sReturnMsg = ia.UpdateReceiptCosts(CStr(vInvoice(i, VINV_PARTNO)), CLng(Trim(lblPon)), _
'                CInt(Trim(vInvoice(i, VINV_ITEMNO))), CStr(vInvoice(i, VINV_ITEMREV)), _
'                cQty, ByVal unitCost, True)
   ia.LoggingEnabled = True
   Dim rdo As ADODB.Recordset
   Dim success As Boolean

   ia.DiagnoseOnly = CBool(chkDiagnose.Value)
   ia.Log "Recosting purchase orders completed between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If ia.DiagnoseOnly Then
      ia.Log "Diagnosing Only.  POs will not be updated"
   End If
   ia.Log ""
   
   'first get count
   Dim fromClause As String
   Dim totalRows As Long
   
'SELECT RTRIM(PIPART) AS PIPART, PINUMBER, PIRELEASE, PIITEM,
'RTRIM(PIREV) AS PIREV, PIAQTY, PIAMT + ISNULL(UnitFreight, 0) AS PIAMT, PIADATE, *
'From POITTABLE
'LEFT JOIN ViitTable ON PINUMBER = VITPO AND PIRELEASE = VITPORELEASE AND PINUMBER <> 0
'AND PIITEM = VITPOITEM AND PIREV = VITPOITEMREV
'LEFT JOIN viewUnitFreightByInvoice ON VITNO = InvoiceNo
'WHERE PIADATE BETWEEN '12/01/07' AND '1/31/08'
'AND ( PITYPE = 15 OR PITYPE = 17 )
'ORDER BY PINUMBER, PIRELEASE, PIITEM, PIREV

   
   fromClause = "FROM POITTABLE" & vbCrLf _
      & "LEFT JOIN ViitTable ON PINUMBER = VITPO AND PIRELEASE = VITPORELEASE AND PINUMBER <> 0" & vbCrLf _
      & "AND PIITEM = VITPOITEM AND PIREV = VITPOITEMREV" & vbCrLf _
      & "LEFT JOIN viewUnitFreightByInvoice ON VITNO = InvoiceNo" & vbCrLf _
      & "AND ( PITYPE = 15 OR PITYPE = 17 )" & vbCrLf _
      & "WHERE PIADATE BETWEEN '" & cmbCompletedFrom & "' AND '" & cmbCompletedThru & "'" & vbCrLf
   sSql = "SELECT COUNT(*) " & fromClause
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      totalRows = rdo.Fields(0)
   End If
   Set rdo = Nothing
   
   sSql = "SELECT RTRIM(PIPART) AS PIPART, PINUMBER, PIRELEASE, PIITEM," & vbCrLf _
      & "RTRIM(PIREV) AS PIREV, PIAQTY, PIAMT + ISNULL(UnitFreight, 0) AS PIAMT, PIADATE" & vbCrLf _
      & fromClause _
      & "ORDER BY PINUMBER, PIRELEASE, PIITEM, PIREV"
      
   Dim totalItems As Integer
   Dim recostedItems As Integer
   
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      'apply constraints on closing
      cmdRecost.enabled = False
      cmdCan.enabled = False
      MouseCursor ccHourglass
      
      Dim max As Long
      If chkDiagnose.Value = vbChecked Then
         max = 99999
      ElseIf IsNumeric(txtMax.Text) Then
         max = CLng(txtMax.Text)
      Else
         max = 99999
      End If
      
      If totalRows > max Then
         totalRows = max
      End If
      
      With rdo
         Do While Not .EOF
            totalItems = totalItems + 1
            ia.PartNumber = !PIPART
            If totalItems <= max Then
               StatusBar1.SimpleText = totalItems & " of " & totalRows & ": " _
                  & " PO " & !PINUMBER & " item " & !PIITEM & !PIREV & " " & !PIADATE & " part " & !PIPART & " completed"
                  
'If !PINUMBER = 65612 Then
'   Debug.Print !PIAMT
'End If
'               ia.Log ""
'               ia.Log "Recosting PO " & !PINUMBER & " item " & !PIITEM & !PIREV & !PIADATE & " part " & !PIPART
               If ia.UpdateReceiptCosts(!PIPART, !PINUMBER, !PIITEM, !PIREV, !PIAQTY, !PIAMT, True) _
                  <> "UpdateReceiptCosts failed" Then
                  recostedItems = recostedItems + 1
               End If
            Else
               ia.Log "Did not recost PO " & !PINUMBER & " item " & !PIITEM & !PIREV
            End If
            .MoveNext
         Loop
      End With
   Else
      MsgBox "There are no PO items received between " & cmbCompletedFrom & " and " & cmbCompletedThru
      Set rdo = Nothing
      Exit Sub
   End If
   Set rdo = Nothing
   Dim sMsg As String
   sMsg = "Recosted " & recostedItems & " of " & totalItems & " items."
   ia.Log ""
   ia.Log sMsg
   StatusBar1.SimpleText = sMsg & "  See log"
   cmdRecost.enabled = True
   cmdCan.enabled = True
   MouseCursor ccArrow
   'MsgBox StatusBar1.SimpleText
   
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4153
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSelPO_Click()
   FillGrid
End Sub

Private Sub cmdVew_Click()
   MouseCursor 13
   On Error GoTo DiaErr1

'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("closedruns")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   SetCrystalAction Me

   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("closedruns")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & sFacility & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ShowGroupTree False
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
'      CheckInvoicing
      GetSettings
      bLotsOn = CheckLotStatus
'      Dim mo As New ClassMO
'      ia.FillComboBoxWithMoParts cmbPrt, "WHERE RUNSTATUS IN ( 'CO', 'CL' )"
'      optDateRange_Click
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   
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
'      .ColAlignment(11) = 2
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "PartNumber"
      .Col = 1
      .Text = "PO Number"
      .Col = 2
      .Text = "Item Num"
      .Col = 3
      .Text = "Rev"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "Est Cost"
      .Col = 6
      .Text = "PI Cost"
      .Col = 7
      .Text = "Vit Cost"
      .Col = 8
      .Text = "Inva UnitCost"
      .Col = 9
      .Text = "Lot UnitCost"
      .Col = 10
      .Text = "Apply"
      
      .ColWidth(0) = 2300
      .ColWidth(1) = 1100
      .ColWidth(2) = 700
      .ColWidth(3) = 700
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 700
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   
   bOnLoad = 1
   
End Sub

Function FillGrid() As Integer
   Dim RdoGrd As ADODB.Recordset
   
   On Error Resume Next
   On Error GoTo DiaErr1
       
   MouseCursor ccHourglass
   Dim cInAmt As Currency
   Dim cLotUnitCost As Currency
   Dim cPIAmt As Currency
   Dim cVitCost As Currency
       
   Dim strComFrom, strComThru As String
   strComFrom = Format(cmbCompletedFrom, "mm/dd/yyyy")
   strComThru = Format(cmbCompletedThru, "mm/dd/yyyy")
   
   sSql = "SELECT PIPART, PINUMBER, PIITEM, PIREV, PIAQTY," & vbCrLf _
               & " PIESTUNIT , (PIAMT + ISNULL(UnitFreight, 0)) as PIAMT" & vbCrLf _
               & ", (VitCost + VITADDERS) as VitCost, INAMT, LotUnitcost" & vbCrLf _
            & " From dbo.InvaTable, dbo.LohdTable, dbo.PoitTable, dbo.ViitTable, viewUnitFreightByInvoice " & vbCrLf _
         & " WHERE ( PITYPE = 15 OR PITYPE = 17 ) AND VITPO = PINUMBER AND " & vbCrLf _
               & " VITPORELEASE = PIRELEASE AND " & vbCrLf _
               & " VITPOITEM = PIITEM AND " & vbCrLf _
               & " VITPOITEMREV = PIREV AND " & vbCrLf _
               & " VITNO = InvoiceNo AND " & vbCrLf _
               & " LOTPO = PINUMBER AND " & vbCrLf _
               & " LOTPOITEM = PIITEM AND " & vbCrLf _
               & " LOTPOITEMREV = PIREV AND " & vbCrLf _
               & " INPONUMBER = PINUMBER AND " & vbCrLf _
               & " INPOITEM = PIITEM AND " & vbCrLf _
               & " INPOREV = PIREV AND " & vbCrLf _
               & " PIADATE Between '" & strComFrom & "' AND '" & strComThru & "' " & vbCrLf _
               & " ORDER BY 1"
               '& " (PIAMT + ISNULL(UnitFreight, 0)) <> INAMT " & vbCrLf _

Debug.Print sSql

   Grd.Rows = 1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
       With RdoGrd
           Do Until .EOF
            
            cInAmt = CDbl(Trim(!INAMT))
            cLotUnitCost = CDbl(Trim(!LotUnitCost))
            cPIAmt = CDbl(Trim(!PIAMT))
            cVitCost = CDbl(Trim(!VitCost))
            
            If ((Round(cInAmt, 2) <> Round(cPIAmt, 2)) Or (Round(cPIAmt, 2) <> Round(cLotUnitCost, 2))) Then
               Grd.Rows = Grd.Rows + 1
               Grd.Row = Grd.Rows - 1
               Grd.Col = 0
               Grd.Text = "" & Trim(!PIPART)
               Grd.Col = 1
               Grd.Text = "" & Trim(!PINUMBER)
               Grd.Col = 2
               Grd.Text = "" & Trim(!PIITEM)
               Grd.Col = 3
               Grd.Text = "" & Trim(!PIREV)
               Grd.Col = 4
               Grd.Text = "" & Trim(!PIAQTY)
               Grd.Col = 5
               Grd.Text = "" & Trim(!PIESTUNIT)
               Grd.Col = 6
               Grd.Text = "" & Trim(!PIAMT)
               Grd.Col = 7
               Grd.Text = "" & Trim(!VitCost)
               Grd.Col = 8
               Grd.Text = "" & Trim(!INAMT)
               Grd.Col = 9
               Grd.Text = "" & Trim(!LotUnitCost)
               Grd.Col = 10
               Set Grd.CellPicture = Chkno.Picture
            End If
            
            .MoveNext
         Loop
      ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   
   MouseCursor ccArrow
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'SaveSettings
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveSettings
   On Error Resume Next
   FormUnload
   Set Close06 = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'cmbCompletedFrom = Format(ES_SYSDATE, "mm/dd/yy")
   'cmbCompletedThru = Format(ES_SYSDATE, "mm/dd/yy")
   'cmbCloseDate = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub


Private Sub cmbCloseDate_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiFina", "Close07.Bits", Trim(str(chkDiagnose)) _
      & CStr(chkRecostPicks.Value) & CStr(chkRecostShipments.Value) & "0000000000000"
   SaveSetting "Esi2000", "EsiFina", "Close07.Max", txtMax.Text
   SaveSetting "Esi2000", "EsiFina", "Close07.From", cmbCompletedFrom.Text
   SaveSetting "Esi2000", "EsiFina", "Close07.Thru", cmbCompletedThru.Text
End Sub

Private Sub GetSettings()
   Dim bits As String
   bits = GetSetting("Esi2000", "EsiFina", "Close07.Bits", "0000000000000000")
   chkDiagnose.Value = CInt(Mid(bits, 1, 1))
   'chkRecostPicks.Value = CInt(Mid(bits, 2, 1))
   chkRecostPicks.Value = 1                        'not currently optional
   'chkRecostShipments.Value = CInt(Mid(bits, 3, 1))
   chkRecostShipments.Value = 0                    'not currently
   txtMax.Text = GetSetting("Esi2000", "EsiFina", "Close07.Max", "99999")
   cmbCompletedFrom.Text = GetSetting("Esi2000", "EsiFina", "Close07.From", Format(ES_SYSDATE, "mm/dd/yy"))
   cmbCompletedThru.Text = GetSetting("Esi2000", "EsiFina", "Close07.Thru", Format(ES_SYSDATE, "mm/dd/yy"))
End Sub


