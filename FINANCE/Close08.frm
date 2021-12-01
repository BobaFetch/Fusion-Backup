VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Close08 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recost Shipped Items"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelMos 
      Caption         =   "Select MO's"
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
      Left            =   5760
      TabIndex        =   18
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   1320
      Width           =   1800
   End
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
      Left            =   11160
      TabIndex        =   17
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   3360
      Width           =   1920
   End
   Begin VB.CheckBox chkDiagnose 
      Caption         =   "Diagnose only (do not update costs)"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2040
      Width           =   3915
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   7035
      Width           =   13245
      _ExtentX        =   23363
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
      Picture         =   "Close08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   360
      Left            =   5400
      Picture         =   "Close08.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print The Report"
      Top             =   240
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "Close08.frx":0938
      Height          =   350
      Left            =   4920
      Picture         =   "Close08.frx":0E12
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   240
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
      Left            =   11160
      TabIndex        =   4
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   2520
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
      Left            =   6240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1200
   End
   Begin VB.Frame fraDateRange 
      Height          =   1155
      Left            =   540
      TabIndex        =   11
      Top             =   720
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
         TabIndex        =   15
         Top             =   1110
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "shipped items"
         Height          =   255
         Left            =   3780
         TabIndex        =   14
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Through"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recost shipments from"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4695
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   2400
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "Close08.frx":12EC
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "Close08.frx":1676
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Recost Shipped Items"
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
      TabIndex        =   9
      Top             =   300
      Width           =   5775
   End
End
Attribute VB_Name = "Close08"
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

Private Sub cmdRecost_Click()
   'Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   Dim iList As Integer
   bCantClose = 0
   
   'get the list of all receipts between the requested dates
   Dim ia As New ClassInventoryActivity
   Dim sReturnMsg As String
   ia.LoggingEnabled = True
   'Dim rdo As ADODB.Recordset
   Dim success As Boolean

   ia.DiagnoseOnly = CBool(chkDiagnose.Value)
   ia.Log "Recosting items shipped between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If ia.DiagnoseOnly Then
      ia.Log "Diagnosing Only.  POs will not be updated"
   End If
   ia.Log ""
   
   Dim totalItems As Integer
   Dim recostedItems As Integer
   
   Dim strPart As String
   Dim strPSNum As String
   Dim strPIItem As String
   Dim strPSCust As String
   Dim strPSDate As String
   recostedItems = 1
   
   For iList = 1 To Grd.Rows - 1
      cmdRecost.enabled = False
      cmdCan.enabled = False
      
      Grd.Col = 9
      Grd.Row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
        
         Grd.Col = 0
         strPart = Trim(Grd.Text)
         Grd.Col = 1
         strPSNum = Trim(Grd.Text)
         Grd.Col = 2
         strPSNum = Trim(Grd.Text)
         Grd.Col = 3
         strPSCust = Trim(Grd.Text)
         Grd.Col = 4
         strPSDate = Trim(Grd.Text)
         

         ia.UpdatePackingSlipCosts strPSNum
         recostedItems = recostedItems + 1
         StatusBar1.SimpleText = " PS " & strPSNum & " " & strPSDate & " " & strPSCust & " completed"

      End If
      Next
   
      Dim sMsg As String
      sMsg = "Recosted " & recostedItems & " packing slips."
      ia.Log ""
      ia.Log sMsg
      StatusBar1.SimpleText = sMsg & "  See log"
      cmdRecost.enabled = True
      cmdCan.enabled = True
      MouseCursor ccArrow
   
   
End Sub

Private Sub old_cmdRecost_Click()
   'Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   
   'get the list of all receipts between the requested dates
   Dim ia As New ClassInventoryActivity
   Dim sReturnMsg As String
   ia.LoggingEnabled = True
   Dim rdo As ADODB.Recordset
   Dim success As Boolean

   ia.DiagnoseOnly = CBool(chkDiagnose.Value)
   ia.Log "Recosting items shipped between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If ia.DiagnoseOnly Then
      ia.Log "Diagnosing Only.  POs will not be updated"
   End If
   ia.Log ""
   
   'first get count
   Dim fromClause As String
   Dim totalRows As Long
   
'SELECT PSNUMBER, PSCUST, PSDATE
'From PshdTable
'Where PSTYPE = 1 And PSPRINTED Is Not Null
'AND PSDATE BETWEEN '1/1/08' AND '1/24/08'
   fromClause = "FROM PshdTable" & vbCrLf _
      & "WHERE PSDATE BETWEEN '" & cmbCompletedFrom & "' AND '" & cmbCompletedThru & "'" & vbCrLf _
      & "AND PSTYPE = 1 AND PSPRINTED IS NOT NULL" & vbCrLf
   sSql = "SELECT COUNT(*) " & fromClause
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      totalRows = rdo.Fields(0)
   End If
   Set rdo = Nothing
   
   sSql = "SELECT PSNUMBER, PSCUST, PSDATE" & vbCrLf _
      & fromClause _
      & "ORDER BY PSNUMBER"
      
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
            'ia.PartNumber = !PIPART    'NOT REQUIRED
            If totalItems <= max Then
               StatusBar1.SimpleText = totalItems & " of " & totalRows & ": " _
                  & " PS " & !PsNumber & " " & !psDate & " " & !PSCUST & " completed"
'               If ia.UpdateReceiptCosts(!PIPART, !PINUMBER, !PIITEM, !PIREV, !PIAQTY, !PIAMT, True) _
'                  <> "UpdateReceiptCosts failed" Then
'                  recostedItems = recostedItems + 1
'               End If

               ia.UpdatePackingSlipCosts !PsNumber
               recostedItems = recostedItems + 1
            Else
               ia.Log "Did not recost PS " & !PsNumber & " " & !psDate & " " & !PSCUST
            End If
            .MoveNext
         Loop
      End With
   Else
      MsgBox "There are no packing slips dated between " & cmbCompletedFrom & " and " & cmbCompletedThru
      Exit Sub
   End If
   Dim sMsg As String
   sMsg = "Recosted " & recostedItems & " of " & totalItems & " packing slips."
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


Private Sub cmdSelMos_Click()
   FillGrid
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 9
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
      Grd.Col = 9
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
End Sub

Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 9
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next

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
   
   sSql = "SELECT INPART, PSNUMBER, INPSITEM, PSCUST, " & vbCrLf _
            & "Convert(varchar(12), PSDATE,101) as psDate, INGLPOSTED," & vbCrLf _
            & " lotNumber , INAMT, LOTUNITCOST" & vbCrLf _
   
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
'      .ColAlignment(11) = 2
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "PartNumber"
      .Col = 1
      .Text = "PS Number"
      .Col = 2
      .Text = "PS Item"
      .Col = 3
      .Text = "PS Cust"
      .Col = 4
      .Text = "PS Date"
      .Col = 5
      .Text = "GL Posted"
      .Col = 6
      .Text = "LotNumber"
      .Col = 7
      .Text = "Inva UnitCost"
      .Col = 8
      .Text = "Lot UnitCost"
      .Col = 9
      .Text = "Apply"
      
      .ColWidth(0) = 2300
      .ColWidth(1) = 1000
      .ColWidth(2) = 700
      .ColWidth(3) = 700
      .ColWidth(4) = 1000
      .ColWidth(5) = 800
      .ColWidth(6) = 1500
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 700
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   
   bOnLoad = 1
   
End Sub


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

Function FillGrid() As Integer
   Dim RdoGrd As ADODB.Recordset
   
   On Error Resume Next
   On Error GoTo DiaErr1
   
   Dim cInAmt As Currency
   Dim cLotUnitCost As Currency
   Dim cPIAmt As Currency
   Dim cVitCost As Currency
       
   MouseCursor ccHourglass
       
   Dim strComFrom, strComThru As String
   strComFrom = Format(cmbCompletedFrom, "mm/dd/yyyy")
   strComThru = Format(cmbCompletedThru, "mm/dd/yyyy")
   
   sSql = "SELECT INPART, PSNUMBER, INPSITEM, PSCUST, " & vbCrLf _
            & "Convert(varchar(12), PSDATE,101) as psDate, INGLPOSTED," & vbCrLf _
            & " lotNumber , INAMT, LOTUNITCOST" & vbCrLf _
            & " From InvaTable, PshdTable, LohdTable" & vbCrLf _
         & " Where lotNumber = INLOTNUMBER" & vbCrLf _
            & " AND INPSNUMBER = PSNUMBER" & vbCrLf _
            & " and LOTORIGINALQTY <> 0" & vbCrLf _
            & " AND PSTYPE = 1 And PSPRINTED Is Not Null" & vbCrLf _
            & " AND PSDATE BETWEEN '" & strComFrom & "' AND '" & strComThru & "'" & vbCrLf _
         & " ORDER BY 1"
            '& " AND INAMT <> LOTUNITCOST" & vbCrLf _

    
   Grd.Rows = 1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
       With RdoGrd
           Do Until .EOF
            cInAmt = CDbl(Trim(!INAMT))
            cLotUnitCost = CDbl(Trim(!LotUnitCost))
            
            If (Round(cInAmt, 2) <> Round(cLotUnitCost, 2)) Then
               Grd.Rows = Grd.Rows + 1
               Grd.Row = Grd.Rows - 1
               Grd.Col = 0
               Grd.Text = "" & Trim(!INPART)
               Grd.Col = 1
               Grd.Text = "" & Trim(!PsNumber)
               Grd.Col = 2
               Grd.Text = "" & Trim(!INPSITEM)
               Grd.Col = 3
               Grd.Text = "" & Trim(!PSCUST)
               Grd.Col = 4
               Grd.Text = "" & Trim(!psDate)
               Grd.Col = 5
               Grd.Text = "" & Trim(!INGLPOSTED)
               Grd.Col = 6
               Grd.Text = "" & Trim(!lotNumber)
               Grd.Col = 7
               Grd.Text = "" & Trim(!INAMT)
               Grd.Col = 8
               Grd.Text = "" & Trim(!LotUnitCost)
               Grd.Col = 9
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
   SaveSetting "Esi2000", "EsiFina", "Close08.Bits", Trim(str(chkDiagnose)) _
      & "000000000000000"
   SaveSetting "Esi2000", "EsiFina", "Close08.Max", txtMax.Text
   SaveSetting "Esi2000", "EsiFina", "Close08.From", cmbCompletedFrom.Text
   SaveSetting "Esi2000", "EsiFina", "Close08.Thru", cmbCompletedThru.Text
End Sub

Private Sub GetSettings()
   Dim bits As String
   bits = GetSetting("Esi2000", "EsiFina", "Close08.Bits", "0000000000000000")
   chkDiagnose.Value = CInt(Mid(bits, 1, 1))
   txtMax.Text = GetSetting("Esi2000", "EsiFina", "Close08.Max", "99999")
   cmbCompletedFrom.Text = GetSetting("Esi2000", "EsiFina", "Close08.From", Format(ES_SYSDATE, "mm/dd/yy"))
   cmbCompletedThru.Text = GetSetting("Esi2000", "EsiFina", "Close08.Thru", Format(ES_SYSDATE, "mm/dd/yy"))
End Sub



