VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close All Completed Manufacturing Orders"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbUnclosed 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   2760
      Width           =   255
   End
   Begin VB.ComboBox cmbCompletedFrom 
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox chkDiagnose 
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3420
      Width           =   495
   End
   Begin VB.ComboBox cmbCompletedThru 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "9999"
      Top             =   3060
      Width           =   495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4860
      Width           =   6285
      _ExtentX        =   11086
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
      Picture         =   "ShopSHf07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   360
      Left            =   5100
      Picture         =   "ShopSHf07a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Print The Report"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkIgnoreExpendables 
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2760
      Width           =   495
   End
   Begin VB.CheckBox chkIgnoreUnpicked 
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2460
      Width           =   495
   End
   Begin VB.CheckBox chkInvoices 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "ShopSHf07a.frx":0938
      Height          =   350
      Left            =   5280
      Picture         =   "ShopSHf07a.frx":0E12
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   2340
      Width           =   360
   End
   Begin VB.ComboBox cmbCloseDate 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCloseMOs 
      Caption         =   "Close MOs"
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
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   4020
      Width           =   1200
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close Form"
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
      Left            =   3000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4020
      Width           =   1200
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4980
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5235
      FormDesignWidth =   6285
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Only"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   5160
      TabIndex        =   26
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "View Unclosed"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   5160
      TabIndex        =   25
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Close MO's completed from"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   23
      Top             =   1140
      Width           =   2175
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnose only (do not close)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   300
      TabIndex        =   22
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3420
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   21
      Top             =   1500
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "MO's"
      Height          =   255
      Left            =   4140
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Close a maximum of"
      Height          =   255
      Left            =   300
      TabIndex        =   19
      Top             =   3090
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Close all completed manufacturing orders meeting the selected criteria.  Yield = MO quantity is assumed."
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
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore part type 5's (Expendables)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   300
      TabIndex        =   14
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore unpicked items"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   300
      TabIndex        =   13
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2460
      Width           =   1995
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Require purchased items to be invoiced"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   74
      Left            =   300
      TabIndex        =   12
      ToolTipText     =   "Test Allocated PO Items For Invoices (System Setting)"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date closed"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   10
      Top             =   1860
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHf07a"
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
Private clsADOTmpConn As ClassFusionADO


Private Sub cmbCompletedFrom_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbCompletedThru_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCloseMOs_Click()
   'Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   Dim sUnclosedMO As String
   
   
   bCantClose = 0
   sJournalID = GetOpenJournal("IJ", Format$(cmbCloseDate, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      bByte = 1
   Else
      If sJournalID = "" Then bByte = 0 Else bByte = 1
   End If
   If bByte = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period " & cmbCloseDate & ".", _
         vbInformation, Caption
      cmdCloseMOs.enabled = True
      Exit Sub
   End If
   
   'get the list of all completed
   Dim mo As New ClassMO
   mo.LoggingEnabled = True
   Dim RdoRun As ADODB.Recordset
   Dim success As Boolean

   mo.DiagnoseOnly = CBool(chkDiagnose.Value)
   mo.Log "Closing manufacturing orders completed between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If mo.DiagnoseOnly Then
      mo.Log "Diagnosing Only.  MOs will not be closed"
      'mo.Log ""
   End If
   
   ' Create new connection for the dataset
   CreateNewDBConn
   
   If (clsADOTmpConn Is Nothing) Then
      MsgBox "Couldn't create second database connection.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   sSql = "select rtrim(RUNREF) as RUNREF, RUNNO, RUNCOMPLETE" & vbCrLf _
      & "from RunsTable" & vbCrLf _
      & "where RUNSTATUS = 'CO' and RUNCOMPLETE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
      & "and RUNCOMPLETE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf _
      & "order by RUNREF, RUNNO"
   bSqlRows = clsADOTmpConn.GetDataSet(sSql, RdoRun, ES_KEYSET)
   
   Dim totalMos As Integer
   Dim closedMos As Integer
   
   If bSqlRows Then
      'apply constraints on closing
      cmdCloseMOs.enabled = False
      cmdCan.enabled = False
      MouseCursor ccHourglass
      mo.CloseRequiresInvoices = CBool(chkInvoices.Value)
      mo.CloseIgnoresUnpicked = CBool(chkIgnoreUnpicked.Value)
      mo.CloseIgnoresExpendables = CBool(chkIgnoreExpendables.Value)
      mo.CloseDate = CDate(cmbCloseDate.Text)
      Dim max As Long
      If chkDiagnose.Value = vbChecked Then
         max = 9999
      ElseIf IsNumeric(txtMax.Text) Then
         max = CLng(txtMax.Text)
      Else
         max = 9999
      End If
         
      With RdoRun
         Do While Not .EOF
            totalMos = totalMos + 1
            mo.PartNumber = !RUNREF
            mo.RunNumber = !Runno
            sUnclosedMO = "Update EsReportClosedRunsLog SET LOG_CLOSED=1 WHERE LOG_PARTNO='" & mo.PartNumber & "' AND LOG_RUNNO = " & mo.RunNumber
            If totalMos <= max Then
               StatusBar1.SimpleText = "Closing # " & totalMos & ": " & mo.PartNumber & " run " & mo.RunNumber _
                  & " completed " & !RUNCOMPLETE
               If mo.CloseMO() Then
                  closedMos = closedMos + 1
                  clsADOCon.ExecuteSQL sUnclosedMO
               End If
            Else
               mo.Log "Did not close MO # " & totalMos & ": " & mo.PartNumber & " run " & mo.RunNumber
            End If
            .MoveNext
         Loop
      End With
      
      Set RdoRun = Nothing
         'clsADOCon.CommitTrans
'      If InStr(1, returnMessage, "failed") = 0 Then
'         CloseMoFinalize = True
'         Log "MO " & SPartRef & " Run " & nRunNo & " Was Closed.  Total cost: " & cRunCost
'      Else
'         clsADOCon.RollbackTrans
'         CloseMoFinalize = False
'         Log "MO " & SPartRef & " Run " & nRunNo & " Closure failed"
'      End If
   
   Else
      MsgBox "There are no open manufacturing orders completed between " & cmbCompletedFrom & " through " & cmbCompletedThru
      cmdCloseMOs.enabled = True
      Set RdoRun = Nothing
      Exit Sub
   End If
   
   ' Close the connection
   CloseDBConn
   
   Dim sMsg As String
   sMsg = "Closed " & closedMos & " of " & totalMos & " completed MOs."
   mo.Log ""
   mo.Log sMsg
   StatusBar1.SimpleText = sMsg & "  See log"
   cmdCloseMOs.enabled = True
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


Private Sub cmdVew_Click()
    MouseCursor 13
    On Error GoTo DiaErr1
    
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   'get custom report name if one has been defined
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("closedruns")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   'pass formulas
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowUnclosedOnly"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add cbUnclosed.Value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'pass Crystal SQL if required
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
    
'
'    'SetMdiReportsize MdiSect
'    MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'    MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'    MdiSect.crw.Formulas(2) = "ShowUnclosedOnly = " & cbUnclosed.Value
'
'
'   sCustomReport = GetCustomReport("closedruns")
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   'SetCrystalAction Me
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
      CheckInvoicing
      GetSettings
      bLotsOn = CheckLotStatus
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
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
   'Set AdoParameter = Nothing
   'Set AdoQry = Nothing
   On Error Resume Next
   FormUnload
   Set ShopSHf07a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub cmbCloseDate_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub cmbCloseDate_LostFocus()
   cmbCloseDate = CheckDate(cmbCloseDate)
   
End Sub

Private Function CheckInvJournal() As Byte
   Dim b As Byte
   sJournalID = GetOpenJournal("IJ", Format$(cmbCloseDate, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period.", _
         vbExclamation, Caption
      CheckInvJournal = 0
      Sleep 500
      Unload Me
   Else
      CheckInvJournal = 1
   End If
   
End Function

Public Sub CheckInvoicing()
   Dim RdoInv As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT COVERIFYINVOICES FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then chkInvoices.Value = RdoInv!COVERIFYINVOICES
   ClearResultSet RdoInv
   Set RdoInv = Nothing

End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiProd", "ShopSHf07a", Trim(str(chkIgnoreUnpicked.Value)) _
      & Trim(str(chkIgnoreExpendables)) & Trim(str(chkDiagnose)) & "00000"
   SaveSetting "Esi2000", "EsiProd", "ShopSHf07aMax", txtMax.Text
   SaveSetting "Esi2000", "EsiProd", "ShopSHf07aMax.From", cmbCompletedFrom.Text
   SaveSetting "Esi2000", "EsiProd", "ShopSHf07aMax.Thru", cmbCompletedThru.Text
   SaveSetting "Esi2000", "EsiProd", "ShopSHf07aMax.Closed", cmbCloseDate.Text
   
End Sub

Private Sub GetSettings()
   Dim bits As String
   bits = GetSetting("Esi2000", "EsiProd", "ShopSHf07a", "00000000")
   If Len(bits) < 6 Then bits = "000000"
   chkIgnoreUnpicked.Value = CInt(Mid(bits, 1, 1))
   chkIgnoreExpendables.Value = CInt(Mid(bits, 2, 1))
   chkDiagnose.Value = CInt(Mid(bits, 3, 1))
   txtMax.Text = GetSetting("Esi2000", "EsiProd", "ShopSHf07aMax", "9999")
   cmbCompletedFrom.Text = GetSetting("Esi2000", "EsiProd", "ShopSHf07aMax.From", Format(ES_SYSDATE, "mm/dd/yy"))
   cmbCompletedThru.Text = GetSetting("Esi2000", "EsiProd", "ShopSHf07aMax.Thru", Format(ES_SYSDATE, "mm/dd/yy"))
   cmbCloseDate.Text = GetSetting("Esi2000", "EsiProd", "ShopSHf07aMax.Closed", Format(ES_SYSDATE, "mm/dd/yy"))
End Sub


Private Function CreateNewDBConn()

   Dim strConStr As String
   Dim strDBName As String
   
   strDBName = sDataBase
   
   If ((sSaAdmin <> "") And (sSaPassword <> "") And _
      (sserver <> "") And (strDBName <> "")) Then
      
      Set clsADOTmpConn = New ClassFusionADO
      
      strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
               sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & strDBName & ";"
      
      'MsgBox strConStr
      Dim ErrNum As Long
      Dim ErrDesc As String
      
      If clsADOTmpConn.OpenConnection(strConStr, ErrNum, ErrDesc) = False Then
        MsgBox "An error occured while trying to connect to the specified database:" & Chr(13) & Chr(13) & _
               "Error Number = " & CStr(ErrNum) & Chr(13) & _
               "Error Description = " & ErrDesc, vbOKOnly + vbExclamation, "  DB Connection Error"
      End If
   End If
End Function

Private Function CloseDBConn()
   If (Not clsADOTmpConn Is Nothing) Then
      Set clsADOTmpConn = Nothing
   End If
End Function
