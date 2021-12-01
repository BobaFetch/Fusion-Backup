VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PackPSf01b
   Caption = "Reprint Inventory Labels"
   ClientHeight = 2580
   ClientLeft = 0
   ClientTop = 0
   ClientWidth = 8445
   Icon = "NewPackPSf01b.frx":0000
   ScaleHeight = 2580
   ScaleWidth = 8445
End
Attribute VB_Name = "PackPSf01b"
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
Dim bGoodItems As Byte
Dim bGoodSO As Byte
Dim bPsSaved As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd





Private Sub cmbPartNumber_Click()
   'if there is a quantity, display it
   
   txtPartQty.Text = ""
   
   If cmbPartNumber.ListCount > 0 Then
      Dim s As String, n As Integer, n2 As Integer
      s = cmbPartNumber.Text
      n = InStr(1, s, "(")
      If n > 0 Then
         n2 = InStr(n, s, ")")
         If n2 > n Then
            txtPartQty.Text = Mid(s, n + 1, n2 - n)
         End If
      End If
   End If
End Sub

Private Sub cmbPsn_Click()
   SelectParts
End Sub


Private Sub SelectParts()
   If optPackingSlip.Value Then
      FillPackingSlipPartsCombo
   Else
      FillSalesOrderPartsCombo
   End If
   
End Sub

Private Sub Form_Activate()
   Dim X As Printer
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      For Each X In Printers
         If Left(X.DeviceName, 9) <> "Rendering" Then _
                 cmbPrinters.AddItem X.DeviceName
      Next
      bOnLoad = 0
      
      PopulateCombos
      
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
   
   Dim i As Integer
   cmbPartNumber.Clear
   
   On Error GoTo DiaErr1
   
   Dim RdoCmb As rdoResultset
   
   '    sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM PshdTable," _
   '        & "PsitTable WHERE (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 AND PSSHIPPED=1) " _
   '        & "AND PSNUMBER=PIPACKSLIP"
   
   sSql = "select PARTNUM, PIPART, PIQTY from psittable" & vbCrLf _
          & "join PartTable on PIPART = PARTREF" & vbCrLf _
          & "where pipackslip = '" & cmbPsn.Text & "'" & vbCrLf _
          & "order by PARTNUM"
   
   bSqlRows = GetDataSet(RdoCmb)
   If bSqlRows Then
      With RdoCmb
         'cmbPartNumber = "" & Trim(!PARTNUM)
         Do Until .EOF
            AddComboStr cmbPartNumber.hwnd, "" & Trim(!PARTNUM) & " (" _
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

Public Sub FillSalesOrderPartsCombo()
   
   Dim i As Integer
   cmbPartNumber.Clear
   
   On Error GoTo DiaErr1
   
   Dim RdoCmb As rdoResultset
   
   sSql = "select PARTNUM, ITQTY from soittable" & vbCrLf _
          & "join PartTable on ITPART = PARTREF" & vbCrLf _
          & "where itso = " & cmbPsn.Text & vbCrLf _
          & "order by PARTNUM"
   
   bSqlRows = GetDataSet(RdoCmb)
   If bSqlRows Then
      With RdoCmb
         'cmbPartNumber = "" & Trim(!PARTNUM)
         Do Until .EOF
            AddComboStr cmbPartNumber.hwnd, "" & Trim(!PARTNUM) & " (" _
               & !ITQTY & ")"
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
   sProcName = "FillSalesOrderPartsCombo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub




Private Sub PopulateCombos()
   cmbPsn.Clear
   If optPackingSlip.Value = True Then
      lblSelect.Caption = "Packing Slip"
      FillPackingSlipCombo
   Else
      lblSelect.Caption = "Sales Order"
      FillSalesOrderCombo
   End If
   DoEvents
End Sub

Public Sub FillPackingSlipCombo()
   
   Dim i As Integer
   cmbPsn.Clear
   
   On Error GoTo DiaErr1
   
   Dim RdoCmb As rdoResultset
   
   sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM PshdTable," _
          & "PsitTable WHERE (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 AND PSSHIPPED=1) " _
          & "AND PSNUMBER=PIPACKSLIP"
   bSqlRows = GetDataSet(RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbPsn = "" & Trim(!PSNUMBER)
         Do Until .EOF
            AddComboStr cmbPsn.hwnd, "" & Trim(!PSNUMBER)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   
   cmbPsn.ListIndex = cmbPsn.ListCount - 1
   'bGoodPs = GetPackslip()
   'GetPSItems
   Exit Sub
   
   DiaErr1:
   sProcName = "FillPackingSlipCombo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillSalesOrderCombo()
   Dim b As Byte
   Dim i As Integer
   Dim RdoSon As rdoResultset
   
   cmbPsn.Clear
   On Error GoTo DiaErr1
   sProcName = "fillcombo"
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO FROM SohdTable,SoitTable " _
          & "WHERE SONUMBER=ITSO AND ITQTY<>0 AND ITPSNUMBER = '' AND ITINVOICE=0 " _
          & "AND ITCANCELED = 0 " _
          & "ORDER BY SONUMBER"
   bSqlRows = GetDataSet(RdoSon)
   If bSqlRows Then
      With RdoSon
         cmbPsn = Format(!SONUMBER, "00000")
         Do Until .EOF
            AddComboStr cmbPsn.hwnd, Format$(!SONUMBER, "00000")
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



Private Sub optPrn_Click()
   PrintLabels
   
End Sub

Private Sub PrintLabels()
   Dim b As Byte
   Dim FormDriver As String
   Dim FormPort As String
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
   CRWLabels.PrinterName = FormPrinter
   CRWLabels.PrinterDriver = FormDriver
   CRWLabels.PrinterPort = FormPort
   CRWLabels.Destination = crptToPrinter
   CRWLabels.ReportFileName = sReportPath & "intcoainvlabel01.rpt"
   CRWLabels.Formulas(0) = "BarCodePart='" & cmbPartNumber & "'"
   CRWLabels.Formulas(1) = "BarCodeQuantity='" & txtPartQty.Text & "'"
   CRWLabels.Formulas(2) = "PartNo='" & cmbPartNumber & "'"
   CRWLabels.Formulas(3) = "Quantity='" & txtPartQty.Text & "'"
   
   'sSql = "{SohdTable.SONUMBER}=" & Val(cmbSon)
   CRWLabels.SelectionFormula = sSql
   CRWLabels.PrinterCopies = Val(txtQty)
   CRWLabels.Action = 1
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optSalesOrders_Click()
   PopulateCombos
End Sub

Private Sub txtQty_LostFocus()
   txtQty = Abs(Val(txtQty))
   If Val(txtQty) < 1 Then txtQty = 1
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty.Text = 1
   
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
