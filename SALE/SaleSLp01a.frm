VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form SaleSLp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Orders (Report)"
   ClientHeight    =   3915
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7185
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "SaleSLp01a.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3915
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optShowTotal 
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdFindEndSO 
      Height          =   375
      Left            =   3600
      Picture         =   "SaleSLp01a.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Find a Sales Order by Customer or PO"
      Top             =   1680
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Options"
      Height          =   855
      Left            =   240
      TabIndex        =   39
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton optRange 
         Caption         =   "Range of Sales Orders"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single Sales Order"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdFindBegSO 
      Height          =   375
      Left            =   3600
      Picture         =   "SaleSLp01a.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Find a Sales Order by Customer or PO"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "SaleSLp01a.frx":0E9E
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp01a.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCpy 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "SaleSLp01a.frx":2392
      Left            =   6480
      List            =   "SaleSLp01a.frx":2394
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   960
      Width           =   585
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5160
      TabIndex        =   30
      Top             =   360
      Width           =   1935
      Begin VB.CommandButton optDis 
         DisabledPicture =   "SaleSLp01a.frx":2396
         DownPicture     =   "SaleSLp01a.frx":3260
         Height          =   330
         Left            =   840
         Picture         =   "SaleSLp01a.frx":412A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   1440
         Picture         =   "SaleSLp01a.frx":4FF4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optCan 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optSvw 
      Caption         =   "SoView"
      Height          =   195
      Left            =   4560
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "SaleSLp01a.frx":5EBE
      Height          =   320
      Left            =   4080
      Picture         =   "SaleSLp01a.frx":6830
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Show Existing Sales Orders"
      Top             =   240
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Number (Contains 300 Max)"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.CheckBox optSln 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optPsn 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optFet 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   6240
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   6240
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optRem 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3915
      FormDesignWidth =   7185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Sales Order Total"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   44
      Top             =   3480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Find by PO"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   43
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Find by PO"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   38
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies "
      Height          =   255
      Index           =   15
      Left            =   5760
      TabIndex        =   34
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   " More Hidden Below"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   1560
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Canceled Items"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   3240
      Width           =   1755
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   25
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Printed Only) - Disabled"
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   21
      Top             =   4920
      Width           =   3345
   End
   Begin VB.Label lblEnd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2280
      TabIndex        =   20
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblBeg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2280
      TabIndex        =   19
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Notes"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip Numbers"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feature Options"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Information"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   2010
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   2010
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   2010
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Sales Order"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order No"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1905
   End
End
Attribute VB_Name = "SaleSLp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Added Show Canceled... Posted notice that it works only
'        with slesh01.rpt (standard)
'03/22/2010 BBS Added logic to lookup by Purchase Order
Option Explicit
Dim bGoodBeg As Byte
Dim bGoodEnd As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetLastSalesOrder()
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT SONUMBER,SOTYPE FROM SohdTable " _
          & " ORDER BY SONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         lblBeg = "" & Trim(!SOTYPE)
         txtBeg = Format$(!SoNumber, SO_NUM_FORMAT)
         lblEnd = "" & Trim(!SOTYPE)
         txtEnd = Format$(!SoNumber, SO_NUM_FORMAT)
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim sCustomReport As String
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   sCustomReport = GetCustomReport("sleco01")
'   If Left$(sCustomReport, 7) = "sleco01" Then
'      optCan.Visible = True
'      z1(11).Visible = True
'   End If
   For b = 1 To 8
      AddComboStr cmbCpy.hWnd, Format$(b, "0")
   Next
   AddComboStr cmbCpy.hWnd, Format$(b, "0")
   cmbCpy = cmbCpy.List(0)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sl01", sOptions)
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optCmt.Value = Val(Mid(sOptions, 2, 1))
      optRem.Value = Val(Mid(sOptions, 3, 1))
      optCom.Value = Val(Mid(sOptions, 4, 1))
      optFet.Value = Val(Mid(sOptions, 5, 1))
      optPsn.Value = Val(Mid(sOptions, 6, 1))
      optSln.Value = Val(Mid(sOptions, 7, 1))
      If Val(Mid(sOptions, 8, 1)) = 2 Then optSingle.Value = True Else optRange.Value = True
      If Len(sOptions) > 8 Then optShowTotal.Value = Val(Mid(sOptions, 9, 1)) Else optShowTotal.Value = 0
   Else
      optShowTotal.Value = 0
   End If
   lblPrinter = GetSetting("Esi2000", "EsiSale", "Psl01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optCmt.Value) _
              & RTrim(optRem.Value) _
              & RTrim(optCom.Value) _
              & RTrim(optFet.Value) _
              & RTrim(optPsn.Value) _
              & RTrim(optSln.Value)
   If optSingle.Value = True Then sOptions = sOptions & "1" Else sOptions = sOptions & "2"
   sOptions = sOptions & RTrim(optShowTotal.Value)
   SaveSetting "Esi2000", "EsiSale", "sl01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiSale", "Psl01", lblPrinter
   
End Sub


'BBS Added this routine on 03/22/2010 for Ticket #28317
'Private Sub cmbPurchOrder_Click()
'    Dim sPONumber As String
'    sPONumber = cmbPurchOrder
'    bGoodBeg = GetSalesOrder(sPONumber)
'End Sub

'Private Sub cmbPurchOrder_LostFocus()
'    Dim sPONumber As String
'    sPONumber = cmbPurchOrder
'    bGoodBeg = GetSalesOrder(sPONumber)
'End Sub

'Private Sub cmbSon_Change()
'   txtBeg = cmbSon
'
'End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdFindBegSO_Click()
    SOLookup.lblControl = "txtBeg"
    SOLookup.lblSONumber = txtBeg.Text
    SOLookup.lblSoType = lblBeg.Caption
    SOLookup.Show
    If optSingle.Value = True Then
        txtEnd = txtBeg
        lblEnd = lblBeg
    End If
End Sub

Private Sub cmdFindEndSO_Click()
    SOLookup.lblControl = "txtEnd"
    SOLookup.lblSONumber = txtEnd.Text
    SOLookup.lblSoType = lblEnd.Caption
    SOLookup.Show
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub cmdVew_Click()
   If cmdVew.Value = True Then
      '        SoTree.Show
      '        cmdVew.Value = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombos
      txtBeg = GetSetting("Esi2000", "EsiSale", "LastRevisedSO", "")
      txtEnd = GetSetting("Esi2000", "EsiSale", "LastRevisedSO", "")
      
      If Val(txtBeg) = 0 Or Val(txtEnd) = 0 Then GetLastSalesOrder
      bGoodBeg = GetSalesOrder(txtBeg, lblBeg)
      bGoodEnd = GetSalesOrder(txtEnd, lblEnd)
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   GetOptions
   optSingle.Value = True
   SetScreenOptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLp01a = Nothing
   
End Sub


Private Sub PrintReport()
   Dim sBegSoNumber, sEndSoNumber As String
   MouseCursor 13
   On Error GoTo DiaErr1
   fraPrn.Enabled = False
   sBegSoNumber = Trim(str(Val(txtBeg)))
   sEndSoNumber = Trim(str(Val(txtEnd)))
   
   ' if ITAR/EAR SOs, alert the user
   Dim rs As ADODB.Recordset
   Dim sos As String
   sSql = "select SONUMBER from SohdTable" & vbCrLf _
      & "where SONUMBER between " & sBegSoNumber & " and " & sEndSoNumber & vbCrLf _
      & "and SOITAREAR = 1 order by SONUMBER"
   If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) <> 0 Then
      With rs
         Do Until rs.EOF
            If Len(sos) > 0 Then
               sos = sos & ","
            End If
            sos = sos & !SoNumber
            .MoveNext
         Loop
      End With
      MsgBox "One or more SOs in ITAR/EAR status: " & sos
   End If
   rs.Close
   Set rs = Nothing
  
   Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    
    aFormulaName.Add "CompanyName"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaName.Add "ShowComments"
    aFormulaValue.Add CStr(optCmt)
    aFormulaName.Add "ShowExDescription"
    aFormulaValue.Add CStr(optExt)
    aFormulaName.Add "ShowRemarks"
    aFormulaValue.Add CStr(optRem)
    aFormulaName.Add "ShowPackingSlipNumbers"
    aFormulaValue.Add CStr(optPsn)
    aFormulaName.Add "ShowTotal"
    aFormulaValue.Add CStr(optShowTotal)
    
    
    sCustomReport = GetCustomReport("sleco01")

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   
   sSql = "{SohdTable.SONUMBER}>=" & sBegSoNumber & " AND {SohdTable.SONUMBER}<=" & sEndSoNumber
   'If Left$(sCustomReport, 7) = "sleco01" Then
      If optCan.Value = vbUnchecked Then _
                        sSql = sSql & " AND {SoitTable.ITCANCELED} = 0"
   'End If
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
   If optPrn Then
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, Val(cmbCpy)
   Else
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
   End If
    
   
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue

   
   MouseCursor 0
   fraPrn.Enabled = True
   Exit Sub
   
DiaErr1:
   fraPrn.Enabled = True
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optCan_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCom_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   txtBeg_LostFocus
   txtEnd_LostFocus
   If bGoodBeg And bGoodEnd Then
        If Val(txtBeg) > Val(txtEnd) Then
            MsgBox "Beginning Sales Order Cannot Be Greater Than Ending Sales Order", vbExclamation, Caption
        Else
            PrintReport
        End If
   Else
      MsgBox "Sales Order Wasn't Found.", vbExclamation, Caption
   End If
   
End Sub

Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFet_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optFet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   txtBeg_LostFocus
   txtEnd_LostFocus
   
   If bGoodBeg And bGoodEnd Then
        If Val(txtBeg) > Val(txtEnd) Then
            MsgBox "Beginning Sales Order Cannot Be Greater Than Ending Sales Order", vbExclamation, Caption
        Else
          PrintReport
        End If
   Else
      MsgBox "Sales Order Wasn't Found.", vbExclamation, Caption
   End If
   
End Sub

Private Sub optPsn_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optPsn_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optRange_Click()
    SetScreenOptions
End Sub

Private Sub optRem_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optRem_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSingle_Click()
    SetScreenOptions
End Sub

Private Sub optSln_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optSln_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub optSvw_Click()
   If optSvw.Value = vbChecked Then
      optSvw.Value = vbUnchecked
      On Error Resume Next
      txtBeg_Click
      txtBeg.SetFocus
   End If
   
End Sub

Private Sub txtBeg_Click()
   bGoodBeg = GetSalesOrder(txtBeg, lblBeg)
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckLen(txtBeg, SO_NUM_SIZE)
   txtBeg = Format(Abs(Val(txtBeg)), SO_NUM_FORMAT)
   bGoodBeg = GetSalesOrder(txtBeg, lblBeg)
   If optSingle.Value = True Then
        txtEnd = txtBeg
        lblEnd = lblBeg
   End If
End Sub

Private Sub txtEnd_Click()
    bGoodEnd = GetSalesOrder(txtEnd, lblEnd)
End Sub

Private Sub txtEnd_LostFocus()
   ' MM If the single option is selected then the end EndPO is same as the begin and the flag.
   If optSingle.Value = True Then
      txtEnd = txtBeg
      bGoodEnd = bGoodBeg
   Else
      txtEnd = CheckLen(txtEnd, SO_NUM_SIZE)
      txtEnd = Format(Abs(Val(txtEnd)), SO_NUM_FORMAT)
   End If
End Sub

Private Function GetSalesOrder(ByVal sSONum As String, ByRef sLabel As String) As Byte
   Dim RdoGet As ADODB.Recordset
   'Dim sWhereClause As String  'BBS Added on 03/22/2010 for Ticket #28317
   
   On Error GoTo DiaErr1
   GetSalesOrder = False
'   If (Val(txtBeg) = 0) And (PONumber = "") Then Exit Function  'BBS Added on 03/22/2010 for Ticket #28317
'   If (PONumber = "") Then sWhereClause = "SONUMBER=" & txtBeg Else sWhereClause = "SOPO='" & PONumber & "'" 'BBS Added on 03/22/2010 for Ticket #28317
   sSql = "SELECT SONUMBER,SOTYPE,SOPO FROM SohdTable WHERE SONUMBER=" & sSONum
   'sWhereClause 'BBS Added on 03/22/2010 for Ticket #28317
 
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            sLabel = Trim(!SOTYPE)
            GetSalesOrder = True
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
   Else
      sLabel = ""
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   GetSalesOrder = False
   
End Function

Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   'Dim sYear As String
   On Error GoTo DiaErr1
   'List = Format(Now, "yyyy")
   'iList = iList - 3
   'sYear = Trim$(iList) & "-" & Format(Now, "mm-dd")
   'sSql = "Qry_FillSalesOrders '" & sYear & "'"
   sSql = "Qry_FillSalesOrders '" & DateAdd("yyyy", -3, Now) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         iList = -1
         Do Until .EOF
            iList = iList + 1
            If iList > 999 Then Exit Do
            AddComboStr txtBeg.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
            AddComboStr txtEnd.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
            'AddComboStr cmbPurchOrder.hwnd, Trim(!SOPO)  'BBS Added on 03/22/2010 for Ticket #28317
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
      Exit Sub
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SetScreenOptions()
    If optSingle.Value = True Then
        txtEnd.Visible = False
        lblEnd.Visible = False
        z1(0).Caption = "Sales Order"
        z1(1).Caption = ""
        cmdFindEndSO.Visible = False
        Label1(1).Visible = False
    Else
        txtEnd.Visible = True
        lblEnd.Visible = True
        z1(0).Caption = "Starting Sales Order"
        z1(1).Caption = "Ending Sales Order"
        cmdFindEndSO.Visible = True
        Label1(1).Visible = True
    End If
    


End Sub
