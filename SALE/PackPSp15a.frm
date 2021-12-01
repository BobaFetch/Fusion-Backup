VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PackPSp15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprint Inventory Labels"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPartQty 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   1260
      Width           =   795
   End
   Begin VB.ComboBox cmbPsn 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Select Packing Slip (limit to 5000 desc) or type in a Packing Slip to repopulate drop-down"
      Top             =   780
      Width           =   2175
   End
   Begin VB.ComboBox cmbPartNumber 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1260
      Width           =   3615
   End
   Begin VB.TextBox txtQty 
      Height          =   288
      Left            =   7740
      TabIndex        =   4
      Tag             =   "2"
      Text            =   "1"
      ToolTipText     =   "Number Of Labels"
      Top             =   1260
      Width           =   372
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   7200
      Picture         =   "PackPSp15a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ComboBox cmbPrinters 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   7800
      Picture         =   "PackPSp15a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   7200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin Crystal.CrystalReport CRWLabels 
      Left            =   960
      Top             =   1740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblPrinter 
      Caption         =   "lblPrinter reqd by SetCrystalAction"
      Height          =   255
      Left            =   3540
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   285
      Index           =   4
      Left            =   5340
      TabIndex        =   12
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label Qty"
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   11
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   780
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Printers"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "PackPSp15a"
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


Dim sLabelPrinter As String
Dim bGoodItems   As Byte
Dim bGoodSO      As Byte
Dim bPsSaved     As Byte
Dim bUserTyping As Byte     'BBS Added this on 03/16/2010 for Ticket #21941


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown()  As New EsiKeyBd



Private Sub cmbPartNumber_Click()
    Dim S As String, n1 As Integer, n2 As Integer
    txtPartQty.Text = ""
    S = cmbPartNumber.Text
    n1 = InStr(1, S, "(")
    n2 = InStr(1, S, ")")
    If n1 > 0 And n2 > n1 + 1 Then
        txtPartQty.Text = Mid(S, n1 + 1, n2 - n1 - 1)
    End If

End Sub


Private Sub cmbPsn_Click()
    SelectParts
End Sub


Private Sub SelectParts()
    'If optPackingSlip.Value Then
        FillPackingSlipPartsCombo
    'Else
    '    FillSalesOrderPartsCombo
    'End If
    
End Sub

Private Sub cmbPsn_DropDown()
    If (bUserTyping = 1) Or (cmbPsn.Text = "") Then FillPackingSlipCombo  'BBS Added this on 03/16/2010 for Ticket #21941
End Sub

Private Sub cmbPsn_KeyPress(KeyAscii As Integer)
    bUserTyping = 1 'BBS Added this on 03/16/2010 for Ticket #21941
End Sub


Private Sub Form_Activate()
    Dim x As Printer
    On Error Resume Next
    MdiSect.BotPanel = Caption
    If bOnLoad Then
        For Each x In Printers
            If Left(x.DeviceName, 9) <> "Rendering" Then _
                cmbPrinters.AddItem x.DeviceName
        Next
        bOnLoad = 0
        
        PopulateCombos
        bUserTyping = 0

    End If
    MouseCursor 0
    
End Sub

Private Sub Form_Load()
    FormLoad Me
    FormatControls
    GetOptions
    bOnLoad = 1
    
End Sub

Private Sub cmdCan_Click()
    Unload Me

End Sub

Private Sub optPackingSlip_Click()
    PopulateCombos
End Sub

Public Sub FillPackingSlipPartsCombo()
   
    Dim I     As Integer
    cmbPartNumber.Clear

    On Error GoTo DiaErr1
    
    Dim RdoCmb As ADODB.Recordset

'    sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM PshdTable," _
'        & "PsitTable WHERE (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 AND PSSHIPPED=1) " _
'        & "AND PSNUMBER=PIPACKSLIP"
        
    sSql = "select PARTNUM, PIPART, PIQTY, PIITNO from psittable" & vbCrLf _
        & "join PartTable on PIPART = PARTREF" & vbCrLf _
        & "where pipackslip = '" & cmbPsn.Text & "'" & vbCrLf _
        & "order by PIITNO, PARTNUM"
        
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
    If bSqlRows Then
        With RdoCmb
            'cmbPartNumber = "" & Trim(!PARTNUM)
            Do Until .EOF
                AddComboStr cmbPartNumber.hWnd, "Item " & !PIITNO & ": " & Trim(!PartNum) & " (" _
                    & !PIQTY & ")"
                .MoveNext
            Loop
            .Cancel
        End With
   End If
    Set RdoCmb = Nothing
    
    If cmbPartNumber.ListCount > 0 Then
        cmbPartNumber.ListIndex = 0
    End If
    
    'bGoodPs = GetPackslip()
    'GetPSItems
    Exit Sub

DiaErr1:
    sProcName = "FillPackingSlipPartsCombo"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub

Private Sub PopulateCombos()
    cmbPsn.Clear
'    If optPackingSlip.Value = True Then
        lblSelect.Caption = "Packing Slip"
        FillPackingSlipCombo
'    Else
'        lblSelect.Caption = "Sales Order"
'        FillSalesOrderCombo
'    End If
    DoEvents
End Sub

Public Sub FillPackingSlipCombo()
    Dim I     As Integer
    Dim sExtraWhere, sPackingSlipSearch As String   'BBS Added this on 03/16/2010 for Ticket #21941
    
    If bUserTyping Then sPackingSlipSearch = cmbPsn.Text Else sPackingSlipSearch = ""  'BBS Added this on 03/16/2010 for Ticket #21941
    cmbPsn.Clear

    On Error GoTo DiaErr1
    
    Dim RdoCmb As ADODB.Recordset

'    sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM PshdTable," _
'        & "PsitTable WHERE (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 AND PSSHIPPED=1) " _
'        & "AND PSNUMBER=PIPACKSLIP"

'BBS Added this on 03/15/2010 for Ticket #21941
    If bUserTyping Then sExtraWhere = " AND PSNUMBER LIKE '" & sPackingSlipSearch & "%'" Else sExtraWhere = ""
    
'BBS Changed this query to limit it to 5000 most recent on 03/16/2010 for Ticket #21941
'    I also made it so that if the user types, it limits the query to a substring of what they entered
    sSql = "SELECT DISTINCT TOP(5000) PSNUMBER,PSTYPE,PIPACKSLIP,PSDATE" & vbCrLf _
      & "FROM PshdTable" & vbCrLf _
      & "JOIN PsitTable ON PSNUMBER=PIPACKSLIP" & vbCrLf _
      & "WHERE PSTYPE=1 AND PSSHIPPRINT=1" & sExtraWhere & vbCrLf _
      & "ORDER BY PSDATE DESC"
     
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
        If bSqlRows Then
            With RdoCmb
                cmbPsn = "" & Trim(!PsNumber)
                Do Until .EOF
                    AddComboStr cmbPsn.hWnd, "" & Trim(!PsNumber)
                    .MoveNext
                Loop
                .Cancel
            End With
       End If
    Set RdoCmb = Nothing
'    If bUserTyping <> 1 And cmbPsn.ListCount > 0 Then cmbPsn.ListIndex = 0
    ' cmbPsn.ListIndex = cmbPsn.ListCount - 1
    'bGoodPs = GetPackslip()
    'GetPSItems
    If bUserTyping = 1 Then bUserTyping = 0  'BBS Added this on 03/16/2010 for Ticket #21941
    Exit Sub

DiaErr1:
    bUserTyping = 0 'BBS Added this on 03/16/2010 for Ticket #21941 to reset this in case it causes the error
    sProcName = "FillPackingSlipCombo"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub

Public Sub FillSalesOrderCombo()
    Dim b       As Byte
    Dim I       As Integer
    Dim RdoSon As ADODB.Recordset
    
    cmbPsn.Clear
    On Error GoTo DiaErr1
    sProcName = "fillcombo"
    sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO FROM SohdTable,SoitTable " _
        & "WHERE SONUMBER=ITSO AND ITQTY<>0 AND ITPSNUMBER = '' AND ITINVOICE=0 " _
        & "AND ITCANCELED = 0 " _
        & "ORDER BY SONUMBER"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
        If bSqlRows Then
            With RdoSon
                cmbPsn = Format(!SoNumber, SO_NUM_FORMAT)
                Do Until .EOF
                    AddComboStr cmbPsn.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
                    .MoveNext
                Loop
                .Cancel
            End With
        End If
    'If cmbPsn.ListCount > 0 Then bGoodSO = GetSalesOrder()
    Set RdoSon = Nothing
    
    cmbPsn.ListIndex = cmbPsn.ListCount - 1
    Exit Sub
    
DiaErr1:
    sProcName = "FillSalesOrderCombo"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me
    
End Sub

Private Sub optDis_Click()
   PrintLabels
End Sub

Private Sub optPrn_Click()
    lblPrinter = cmbPrinters
    PrintLabels

End Sub

Private Sub PrintLabels()
   Dim b           As Byte
   Dim FormDriver  As String
   Dim FormPort    As String
   Dim FormPrinter As String
   
   MouseCursor 13
   FormPrinter = Trim(sLabelPrinter)
   If Len(Trim(FormPrinter)) > 0 Then
      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
   Else
      FormPrinter = ""
      FormDriver = ""
      FormPort = ""
   End If
   
   'separate item number and part number:  item nn: part (qty)
   Dim str As String, part As String
   Dim Item As Integer, n1 As Integer, n2 As Integer, n3 As Integer
   str = cmbPartNumber
   n1 = InStr(1, str, " ")
   n2 = InStr(n1, str, ":")
   n3 = InStr(n2, str, "(")
   Item = CInt(Mid(str, n1 + 1, n2 - n1 - 1))
   part = Trim(Mid(str, n2 + 1, n3 - n2 - 1))
   
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
 
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps20.rpt")
 
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps20.rpt"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "PartNumber"
    aFormulaName.Add "PartRef"
    aFormulaName.Add "Quantity"
    aFormulaName.Add "PackSlipNumber"
    aFormulaName.Add "PackSlipItem"

    aFormulaValue.Add CStr("'" & CStr(part) & "'")
    aFormulaValue.Add CStr("'" & CStr(Compress(part)) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtPartQty.Text) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbPsn) & "'")
    aFormulaValue.Add CStr("'" & CStr(Item) & "'")

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, Val(txtQty)

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


Private Sub txtQty_LostFocus()
    txtQty = Abs(Val(txtQty))
    If Val(txtQty) < 1 Then txtQty = 1
    
End Sub

Private Sub FormatControls()
    Dim b As Byte
    b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
    txtQty = 1
    
End Sub


Private Sub GetOptions()
    Dim sOptions As String
    On Error Resume Next
    sLabelPrinter = GetSetting("Esi2000", "System", "Label Printer", sLabelPrinter)
    cmbPrinters = sLabelPrinter
    
End Sub


Private Sub Form_Resize()
    Refresh

End Sub
