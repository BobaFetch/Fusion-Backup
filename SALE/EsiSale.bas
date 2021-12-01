Attribute VB_Name = "EsiSle"
'*** ES/2000 (ES/2001 - ES/2007) is the property of ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/2/03 Customer Permissions
'9/27/03 Commissions
'10/2/03 Prepackaging
'3/29/05 Removed references to JetDb
'4/14/05 Added FillPackingSlip Queries
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'1/12/06 Completed renaming dialogs to be consistent with Fina
'4/25/06 Stored Procedures (See UpdateTables)
'5/16/06 Delete Triggers on SohdTable/PshdTable
'6/1/06 BuildKeys
'6/12/06 ConvertSalesColumns
'6/26/06 Removed Threed32.OCX
'8/10/06 Removed SSTab32.OCX
'9/16/06 Added GetPriceBook (common for Customer and Sales Order)
'1/11/07 Added GetThisCustomer for reports 7.2.7
Option Explicit
'Sales Module code
Public Y As Byte
Public bFoundPart As Byte
Public sCurrForm As String
Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String
Public sPassedPart As String


'Column updates

'Type tUser
'    Adduser As Integer
'    Level   As Integer
'    Group1  As Integer
'    Group2  As Integer
'    Group3  As Integer
'    Group4  As Integer
'    Group5  As Integer
'    Group6  As Integer
'    Group7  As Integer
'    Group8  As Integer
'End Type
'Public User As tUser
'
'9/11/06 Add ins for PROPLA

Public Sub GetPriceBook(frm As Form)
   Dim RdoBook As ADODB.Recordset
   
   On Error GoTo modErr1
   If frm.txtBook.Visible = False Then Exit Sub
   frm.txtBook = "No Price Book"
   frm.BookExpires = ""
   frm.BookExpires.ToolTipText = "The Price Book Expiration Date"
   sSql = "SELECT CUREF,CUPRICEBOOK FROM CustTable WHERE " _
          & "CUREF='" & Compress(frm.cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBook, ES_FORWARD)
   If bSqlRows Then
      With RdoBook
         frm.txtBook = "" & Trim(!CUPRICEBOOK)
         ClearResultSet RdoBook
      End With
   End If
   If frm.txtBook <> "" Then
      On Error Resume Next
      sSql = "SELECT PBHREF,PBHENDDATE FROM PbhdTable WHERE PBHREF='" _
             & Compress(frm.txtBook) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBook, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoBook!PBHENDDATE) Then _
                       frm.BookExpires = Format(RdoBook!PBHENDDATE, "mm/dd/yy")
         If IsDate(frm.BookExpires) Then
            If Format(frm.BookExpires, "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
               frm.txtBook.ForeColor = ES_RED
               frm.BookExpires.ForeColor = ES_RED
               frm.BookExpires.ToolTipText = "The Price Book Has Expired"
            Else
               frm.txtBook.ForeColor = vbBlack
               frm.BookExpires.ForeColor = vbBlack
               frm.BookExpires.ToolTipText = "The Price Book Expiration Date"
            End If
         End If
      Else
         frm.txtBook.ForeColor = vbBlack
         frm.BookExpires.ForeColor = vbBlack
         frm.BookExpires = ""
         frm.BookExpires.ToolTipText = "The Price Book Expiration Date"
      End If
   Else
      frm.txtBook.ForeColor = vbBlack
      frm.BookExpires.ForeColor = vbBlack
      'frm.txtBook = "No Price Book Assigned"
   End If
   Set RdoBook = Nothing
   Exit Sub
   
modErr1:
   sProcName = "getpricebook"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub


Public Function GetQNMConversion() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT QueueMoveConversion FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         If Not IsNull(!QueueMoveConversion) Then
            GetQNMConversion = !QueueMoveConversion
         Else
            GetQNMConversion = 24
         End If
         .Cancel
      End With
      ClearResultSet RdoGet
   Else
      GetQNMConversion = 24
   End If
   Set RdoGet = Nothing
   Exit Function
   
modErr1:
   GetQNMConversion = 24
   
End Function


Public Function CheckCustomerPO() As Byte
   On Error GoTo modErr1
   Dim RdoCpo As ADODB.Recordset
   If Trim(MdiSect.ActiveForm.txtCpo) = "" Then
      CheckCustomerPO = 0
   Else
      sSql = "Qry_GetCustomerPo '" & Compress(MdiSect.ActiveForm.cmbCst) _
             & "','" & Trim(MdiSect.ActiveForm.txtCpo) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpo, ES_FORWARD)
      If bSqlRows Then
         With RdoCpo
            CheckCustomerPO = 1
            ClearResultSet RdoCpo
         End With
      End If
   End If
   Set RdoCpo = Nothing
   Exit Function
   
modErr1:
   sProcName = "CheckCusto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   CheckCustomerPO = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Sub GetLastSalesOrder(sOldSo As String, sNewSo As String, bFillText As Boolean)
   Dim RdoSon As ADODB.Recordset
   Dim lSales As Long
   On Error GoTo DiaErr1
   sSql = "SELECT COLASTSALESORDER From ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
'      If Val(Right$(RdoSon!COLASTSALESORDER, 6)) > 0 Then _
'             MdiSect.ActiveForm.lblLst = Trim(RdoSon!COLASTSALESORDER)
   
      If Val(Right$(RdoSon!COLASTSALESORDER, 6)) > 0 Then
             'MdiSect.ActiveForm.lblLst = Trim(RdoSon!COLASTSALESORDER)
             MdiSect.ActiveForm.lblLst = Left$(RdoSon!COLASTSALESORDER, 1) & Format(Val(Right$(RdoSon!COLASTSALESORDER, 6)), SO_NUM_FORMAT)
      End If
   
   End If
   If Trim(MdiSect.ActiveForm.lblLst) = "" Then
      sSql = "SELECT MAX(SONUMBER)AS SalesOrder,SOTYPE FROM SohdTable GROUP BY SOTYPE "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
      If bSqlRows Then
         With RdoSon
            If Not IsNull(.Fields(0)) Then
               MdiSect.ActiveForm.lblLst = "" & Trim(!SOTYPE) & Format$(!SalesOrder, SO_NUM_FORMAT)
            Else
               MdiSect.ActiveForm.lblLst = "S000000"
            End If
            ClearResultSet RdoSon
         End With
      End If
   End If
   sNewSo = Format(Val(Right(MdiSect.ActiveForm.lblLst, SO_NUM_SIZE)) + 1, SO_NUM_FORMAT)
   If bFillText Then MdiSect.ActiveForm.txtSon = sNewSo
   sOldSo = MdiSect.ActiveForm.lblLst
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub GetNewSalesOrder(ByRef sNewSo As String, ByVal sSoType As String)
   Dim RdoSon As ADODB.Recordset
   Dim lSales As Long
   On Error GoTo DiaErr1
   
      ' Not needed - as the MaxSO# is the next SO.
   sSql = "SELECT (MAX(SONUMBER)+ 1)AS SalesOrder FROM SohdTable" 'WHERE SOTYPE = '" & sSoType & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         If Not IsNull(.Fields(0)) Then
            sNewSo = "" & Format$(!SalesOrder, SO_NUM_FORMAT)
         Else
            sNewSo = SO_NUM_FORMAT
         End If
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetNewSalesOrder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Function GetExistSO(ByVal strBuyerNumber As String, _
            ByRef strSoNum As String, ByRef strNextItem As String) As Boolean
   Dim RdoSon As ADODB.Recordset
   Dim lSales As Long
   On Error GoTo DiaErr1
   
   GetExistSO = True
   
   sSql = "SELECT ITSO, (MAX(ITNUMBER)+ 1)AS Item FROM SoitTable, SohdTable " _
         & " WHERE SoitTable.ITSO = SohdTable.SONUMBER AND SOPO = '" & strBuyerNumber & "' GROUP BY ITSO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         If Not IsNull(.Fields(0)) Then
            strSoNum = "" & Format$(.Fields(0), SO_NUM_FORMAT)
            strNextItem = "" & .Fields(1)
         Else
            GetExistSO = False
         End If
         ClearResultSet RdoSon
      End With
   Else
      GetExistSO = False
   End If
   
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetNewSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function CheckOfExistingSO(strPONumber As String, strPartID As String, ByRef strSoNum As String) As Boolean
   Dim RdoSO As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT ISNULL(MAX(SONUMBER),0) SONUMBER  FROM sohdTable,SoitTable WHERE " _
             & " SONUMBER = ITSO AND SOPO = '" & strPONumber & "'" _
             & "  AND ITPART = '" & Compress(strPartID) & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
   If bSqlRows Then
      With RdoSO
         If (Trim(!SoNumber) = 0) Then
            strSoNum = ""
            CheckOfExistingSO = False
         Else
            strSoNum = Trim(!SoNumber)
            CheckOfExistingSO = True
         End If
         ClearResultSet RdoSO
      End With
   Else
      strSoNum = ""
      CheckOfExistingSO = False
      
   End If
   
   Set RdoSO = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CheckOfExistingSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   

End Function


Public Function GetOldSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   GetOldSalesOrder = 0
   sSql = "SELECT SONUMBER FROM SohdTable WHERE SONUMBER=" _
          & MdiSect.ActiveForm.txtSon & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         GetOldSalesOrder = 1
         ClearResultSet RdoSon
      End With
   Else
      GetOldSalesOrder = 0
   End If
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getoldsal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   GetOldSalesOrder = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function

'Changed to controls 11/6/04

Public Function FindVendor(ContrlCombo As Control, ControlLabel As Control) As Byte
   Dim RdoVed As ADODB.Recordset
   If Len(ContrlCombo) = 0 Then Exit Function
   On Error GoTo modErr1
   sSql = "Qry_GetVendorBasics '" & Compress(ContrlCombo) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         ContrlCombo = "" & Trim(!VENICKNAME)
         ControlLabel = "" & Trim(!VEBNAME)
         FindVendor = 1
         ClearResultSet RdoVed
      End With
   Else
      On Error Resume Next
      ContrlCombo = ""
      ControlLabel = "No Valid Vendor Selected."
      FindVendor = 0
   End If
   Set RdoVed = Nothing
   Exit Function
   
modErr1:
   sProcName = "findvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   FindVendor = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function



'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MdiSect.lblBotPanel.Caption = MdiSect.Caption
   ' TODO: we need to find the logic
'   If Forms.Count < 3 Then
'      iResultSets = RdoCon.ADODB.Recordsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         RdoCon.ADODB.Recordsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Book"
            zGr3Book.Show
         Case "Ordr"
            zGr1Sale.Show
         Case "Pack"
            zGr2Pack.Show
         Case "Comm"
            zGr4Comm.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

Sub Main()
   Dim sAppTitle As String
   If App.PrevInstance Then
      On Error Resume Next
      sAppTitle = App.Title
      App.Title = "%EseSl1"
      SysMsgBox.Width = 4100
      SysMsgBox.msg.Width = 3500
      SysMsgBox.tmr1.Enabled = True
      SysMsgBox.msg = sAppTitle & " Is Already Open."
      SysMsgBox.Show
      Sleep 5000
      AppActivate sAppTitle
   End
   Exit Sub
End If
' Set the Module name before loading the form
sProgName = "Sales"
MainLoad "sale"
GetFavorites "EsiSale"
' save the setting in registry for the module
SetRegistryAppTitle ("EsiSale")
' MM 9/10/2009
'sProgName = "Sales"
MdiSect.Show

End Sub


Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   Select Case sSelected
      Case "Customers"
         cUR.CurrentGroup = "Ordr"
         SaleSLe03a.Show
      Case "Customer List"
         cUR.CurrentGroup = "Ordr"
         SaleSLp02a.Show
      Case "Customer Directory"
         cUR.CurrentGroup = "Ordr"
         SaleSLp03a.Show
      Case "Revise Sales Order"
         cUR.CurrentGroup = "Ordr"
         SaleSLe02a.Show
      Case "New Sales Order"
         cUR.CurrentGroup = "Ordr"
         SaleSLe01a.Show
      Case "Sales Orders (Report)"
         cUR.CurrentGroup = "Ordr"
         SaleSLp01a.Show
      Case "Cancel A Sales Order"
         cUR.CurrentGroup = "Ordr"
         SaleSLf01a.Show
      Case "Sales Order Acknowledgements"
         cUR.CurrentGroup = "Ordr"
         SaleSLp42a.Show
      Case "Customer Sales Order List"
         cUR.CurrentGroup = "Ordr"
         SaleSLp04a.Show
      Case "Sales Order Register"
         cUR.CurrentGroup = "Ordr"
         SaleSLp05a.Show
      Case "Customer List By Zip Code"
         cUR.CurrentGroup = "Ordr"
         SaleSLp06a.Show
      Case "Customer Sales Order List By PO"
         cUR.CurrentGroup = "Ordr"
         SaleSLp07a.Show
      Case "Copy A Sales Order"
         cUR.CurrentGroup = "Ordr"
         SaleSLf02a.Show
      Case "Change A Sales Order Number"
         cUR.CurrentGroup = "Ordr"
         SaleSLf05a.Show
      Case "Revise Booked Dates"
         cUR.CurrentGroup = "Ordr"
         SaleSLf06a.Show
      Case "New Packing Slip"
         cUR.CurrentGroup = "Pack"
         PackPSe01a.Show
      Case "Revise A Packing Slip"
         cUR.CurrentGroup = "Pack"
         PackPSe02a.Show
      Case "Packing Slips"
         cUR.CurrentGroup = "Pack"
         PackPSp01a.Show
      Case "Packing Slip Log"
         cUR.CurrentGroup = "Pack"
         PackPSp02a.Show
      Case "Packing Slip Edit"
         cUR.CurrentGroup = "Pack"
         PackPSp03a.Show
      Case "Cancel A Packing Slip Printing"
         cUR.CurrentGroup = "Pack"
         PackPSf02a.Show
      Case "Cancel A Packing Slip"
         cUR.CurrentGroup = "Pack"
         PackPSf01a.Show
      Case "Bookings By Order Date"
         cUR.CurrentGroup = "Book"
         BookBKp01a.Show
      Case "Bookings By Part Number"
         cUR.CurrentGroup = "Book"
         BookBKp02a.Show
      Case "Bookings By Salesperson"
         cUR.CurrentGroup = "Book"
         BookBKp03a.Show
      Case "Bookings By Division, Class And Code"
         cUR.CurrentGroup = "Book"
         BookBKp04a.Show
      Case "Bookings By Business Unit And Code"
         cUR.CurrentGroup = "Book"
         BookBKp05a.Show
      Case "Bookings By Customer"
         cUR.CurrentGroup = "Book"
         BookBKp06a.Show
      Case "Backlog By Part With Prices"
         cUR.CurrentGroup = "Book"
         BookBLp01a.Show
      Case "Backlog By Salesperson"
         cUR.CurrentGroup = "Book"
         BookBLp04a.Show
      Case "Backlog By Customer"
         cUR.CurrentGroup = "Book"
         BookBLp05a.Show
      Case "Backlog By Scheduled Date"
         cUR.CurrentGroup = "Book"
         BookBLp02a.Show
      Case "Backlog By Sales Order Date"
         cUR.CurrentGroup = "Book"
         BookBLp03a.Show
      Case "Pack Slips Printed Not Invoiced"
         cUR.CurrentGroup = "Pack"
         PackPSp07a.Show
      Case "Pack Slips Not Printed"
         cUR.CurrentGroup = "Pack"
         PackPSp06a.Show
      Case "Delete A Customer"
         cUR.CurrentGroup = "Ordr"
         SaleSLf03a.Show
      Case "Change A Customer Nickname"
         cUR.CurrentGroup = "Ordr"
         SaleSLf04a.Show
      Case "Manufacturing Order Sales Order Allocations"
         cUR.CurrentGroup = "Ordr"
         ShopSHp12a.Show
      Case "Sales Order Allocations By Customer"
         cUR.CurrentGroup = "Ordr"
         ShopSHp11a.Show
      Case "Price Books"
         cUR.CurrentGroup = "Ordr"
         SaleSLe04a.Show
      Case "Delete A Price Book"
         cUR.CurrentGroup = "Ordr"
         SaleSLf09a.Show
      Case "Change A Price Book ID"
         cUR.CurrentGroup = "Ordr"
         SaleSLf07a.Show
      Case "Price Books Report"
         cUR.CurrentGroup = "Ordr"
         SaleSLp08a.Show
      Case "Price Books By Customer"
         cUR.CurrentGroup = "Ordr"
         SaleSLp09a.Show
      Case "Price Books By Part Number"
         cUR.CurrentGroup = "Ordr"
         SaleSLp10a.Show
      Case "Copy A Price Book ID"
         cUR.CurrentGroup = "Ordr"
         SaleSLf08a.Show
      Case "Packing Slip Pick List"
         cUR.CurrentGroup = "Pack"
         PackPSp08a.Show
      Case "Revise Sales Order Commissions"
         cUR.CurrentGroup = "Comm"
         CommCOe02a.Show
      Case "Salespersons"
         cUR.CurrentGroup = "Comm"
         CommCOe01a.Show
      Case "Commission AP Invoice"
         cUR.CurrentGroup = "Comm"
         CommCOe03a.Show
      Case "Commission Status (Report)"
         cUR.CurrentGroup = "Comm"
         CommCOp01a.Show
      Case "Packing Slips Printed Not Shipped"
         cUR.CurrentGroup = "Pack"
         PackPSp09a.Show
      Case "Ship Packaged Goods"
         cUR.CurrentGroup = "Pack"
         PackPSe03a.Show
      Case "Part Availability For A Customer"
         cUR.CurrentGroup = "Pack"
         Intavl03.Show
      Case "Add An Item To A Printed Packing Slip"
         cUR.CurrentGroup = "Pack"
         PackPSe04a.Show
      Case "Backlog By Part Number, Current Month"
         cUR.CurrentGroup = "Book"
         BookBLp06a.Show
      Case "Backlog, 12 Month By Division"
         cUR.CurrentGroup = "Book"
         BookBLp07a.Show
      Case "Shipped Items By Part Number"
         cUR.CurrentGroup = "Pack"
         PackPSp04a.Show
      Case "Shipped Items By Shipping Date"
         cUR.CurrentGroup = "Pack"
         PackPSp05a.Show
      Case "Shipped Items By Sales Person"
         cUR.CurrentGroup = "Pack"
         PackPSp10a.Show
      Case "Shipped Items By Customer"
         cUR.CurrentGroup = "Pack"
         PackPSp11a.Show
      Case "Revise A Packing Slip - Not Shipped"
         cUR.CurrentGroup = "Pack"
         PackPSe05a.Show
      Case "Cancel A Packing Slip Shipped Flag"
         cUR.CurrentGroup = "Pack"
         PackPSf03a.Show
      Case "Cancel A Packing Slip Item (Not Printed)"
         cUR.CurrentGroup = "Pack"
         PackPSf04a.Show
      Case "Cancel A Packing Slip Item (Printed)"
         cUR.CurrentGroup = "Pack"
         PackPSf05a.Show
      Case "Lock/Unlock Sales Orders"
         cUR.CurrentGroup = "Ordr"
         SaleSLf10a.Show
      Case "Split A Packing Slip Item"
         cUR.CurrentGroup = "Pack"
         PackPSe06a.Show
      Case "Revise Request Date For Unshipped PS Items"
         cUR.CurrentGroup = "Pack"
         PackPSe07a.Show
      Case "Company Transfers (Report)"
         cUR.CurrentGroup = "Pack"
         PackPSp13a.Show
      Case "Cancel A Company Transfer"
         cUR.CurrentGroup = "Pack"
         PackPSf06a.Show
      Case "Unshipped Packing Slips By SO Request Date"
         cUR.CurrentGroup = "Pack"
         PackPSp12a.Show
      Case "Part Availability Report"
         BookBKp18a.Show
      Case "Estimate Summary By Customer"
         EstiESp02a.Show
      Case "Create Sales Orders From Estimates"
         SaleSLf11a.Show
      Case "Import Sales Orders"
         SaleSLf12a.Show
      Case "Customer Sales Analysis"
         SaleSLp11a.Show
      Case "Customer Shipment Analysis"
         SaleSLp12a.Show
      Case "Customer Backlog Analysis"
         BookBLp08a.Show
      Case "Inventory Available To Ship"
         PackPSp14a.Show
      Case Else
         MouseCursor 0
   End Select
   On Error GoTo 0
   Exit Sub
   
OpenFavErr1:
   Resume OpenFavErr2
OpenFavErr2:
   MsgBox "ActiveX Error. Can't Load Form..", 48, "System    "
   On Error GoTo 0
   
End Sub

Public Sub FillParts()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillSortedParts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbPrt.hWnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


'Create Tables, etc, here


Public Function AllowPsPrepackaging() As Byte
   Dim RdoPsl As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT COALLOWPSPREPICKS FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_FORWARD)
   If bSqlRows Then
      With RdoPsl
         If Not IsNull(.Fields(0)) Then
            AllowPsPrepackaging = .Fields(0)
         Else
            AllowPsPrepackaging = 0
         End If
         ClearResultSet RdoPsl
      End With
   Else
      AllowPsPrepackaging = 0
   End If
   Set RdoPsl = Nothing
   Exit Function
modErr1:
   AllowPsPrepackaging = 0
   
End Function

'11/4/03 for SO Item Splits
'Use Local Errors
Private Function GetNextAlphaNumericChar(inChar As String) As String
    Dim iNewChar As Integer
    
    If Asc(inChar) = 57 Then
        iNewChar = Asc(inChar) + 8
        GetNextAlphaNumericChar = Trim(Chr$(iNewChar))
    ElseIf Asc(inChar) = 90 Then
        iNewChar = 49
        GetNextAlphaNumericChar = Trim(Chr$(90) & Trim(Chr$(49)))
    Else
        iNewChar = Asc(inChar) + 1
        GetNextAlphaNumericChar = Trim(Chr$(iNewChar))
    End If
    'GetNextAlphaNumericChar = Trim(Chr$(iNewChar))
End Function

Private Function GetNextAlphaChar(inChar As String) As String
    Dim iNewChar As Integer
    
    If Asc(inChar) = 90 Then
        iNewChar = 65
    Else
        iNewChar = Asc(inChar) + 1
    End If
    GetNextAlphaChar = Trim(Chr$(iNewChar))
End Function

Public Function GetNextSORevision(ItemRev As String) As String
    Dim sTemp As String
    'MM not needed should start from Z1 - due to sorting issue.
'    If Len(ItemRev) = 1 And ItemRev = "Z" Then
'        GetNextSORevision = "A1"
'        Exit Function
'    End If
    If Len(ItemRev) = 0 Then
      GetNextSORevision = "A"
   ElseIf Len(ItemRev) = 1 Then
      'More than 26?
      GetNextSORevision = GetNextAlphaNumericChar(Left(ItemRev, 1))
      
   ElseIf Len(ItemRev) = 2 Then
      'More than 26?
      If (ItemRev = "ZZ") Then
         GetNextSORevision = "ZAA"
      Else
         GetNextSORevision = Left(ItemRev, 1) & GetNextAlphaNumericChar(Right(ItemRev, 1))
      End If
      
   ElseIf Len(ItemRev) > 2 Then
         If (Right(ItemRev, 1) = "Z") Then
            sTemp = Left(ItemRev, 1) & GetNextAlphaChar(Mid(ItemRev, 2, 1)) & "A"
         Else
            sTemp = Left(ItemRev, 2) & GetNextAlphaChar(Right(ItemRev, 1))
         End If
         GetNextSORevision = sTemp
         
   Else
      sTemp = Left(ItemRev, 1) & GetNextAlphaNumericChar(Right(ItemRev, 1))
      If Right(sTemp, 1) = "1" Then sTemp = GetNextAlphaNumericChar(Left(sTemp, 1)) & Right(sTemp, 1)
      GetNextSORevision = sTemp
   End If
End Function


'Sales Persons
'SprsTable
'SpcoTable


'1/11/07 Added for reports

Public Sub GetThisCustomer(Optional ControlIsTextBox As Byte)
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   If ControlIsTextBox = 1 Then
      sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
             & Compress(MdiSect.ActiveForm.txtCst) & "'"
   Else
      sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
             & Compress(MdiSect.ActiveForm.cmbCst) & "'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MdiSect.ActiveForm.lblCUName = "" & Trim(RdoRpt!CUNAME)
      ClearResultSet RdoRpt
   Else
      MdiSect.ActiveForm.lblCUName = "*** A Range Of Customers Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

Private Function TwoCharacterSORevsExist(SoNumber As String, ItemNo As String)
    Dim rdoSORev As ADODB.Recordset
    On Error Resume Next
    TwoCharacterSORevsExist = False
    sSql = "SELECT TOP 1 ITREV FROM SoitTable WHERE ITSO=" & SoNumber & " AND ITNUMBER=" & ItemNo & " AND LEN(ITREV)>1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoSORev, ES_FORWARD)
    If bSqlRows Then TwoCharacterSORevsExist = True
    Set rdoSORev = Nothing
End Function

Public Function GetLastRevForSalesOrderItem(SoNumber As String, ItemNo As String)
    Dim rdoSORev As ADODB.Recordset
    Dim TempSql As String
    On Error GoTo modErr1
    
   ' MM 5/16/10 changed the return function name - compiler error.
    TempSql = "SELECT ITREV FROM SoitTable WHERE ITSO=" & SoNumber & " AND ITNUMBER=" & ItemNo
    If TwoCharacterSORevsExist(SoNumber, ItemNo) Then TempSql = TempSql & " AND LEN(ITREV)>1 "
    TempSql = TempSql & " ORDER BY ITREV DESC"
    sSql = TempSql
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoSORev, ES_FORWARD)
    
    Debug.Print sSql
    If bSqlRows Then
        GetLastRevForSalesOrderItem = "" & Trim(rdoSORev!itrev)
    Else
        GetLastRevForSalesOrderItem = ""
    End If
    Set rdoSORev = Nothing
    Exit Function
    
modErr1:
    On Error GoTo 0
    
End Function


Public Sub ResetLastSalesOrderNumber(strPrefix As String)
    On Error Resume Next
    
    Dim rdoSORev As ADODB.Recordset
    On Error GoTo modErr1
    
    sSql = "SELECT CONVERT(CHAR,MAX(SONUMBER)) as MaxSo FROM SohdTable"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoSORev, ES_FORWARD)
    
    Debug.Print sSql
    Dim strLastSO As String
    If bSqlRows Then
      strLastSO = strPrefix & "" & Format(Trim(rdoSORev!MaxSo), SO_NUM_FORMAT)
      sSql = "Update ComnTable SET COLASTSALESORDER = '" & strLastSO & "' Where COREF = 1"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
    Else
      sSql = "Update ComnTable SET COLASTSALESORDER = (SELECT CONVERT(CHAR,MAX(SONUMBER)) FROM SohdTable) Where COREF = 1"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
    End If
    
    Set rdoSORev = Nothing
    
    Exit Sub
    
modErr1:
    On Error GoTo 0

End Sub

Public Function CheckCusCreditLmt(strCust As String)
   Dim RdoCoCrdLmt As ADODB.Recordset
   Dim RdoCustCL As ADODB.Recordset
   
   Dim cComCrdLmt As Currency
   Dim cCusCrdLmt As Currency
   
   On Error Resume Next
   
   If ColumnExists("ComnTable", "CUWARNCREDITLMT") Then
      
      sSql = "SELECT CUWARNCREDITLMT FROM ComnTable WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCoCrdLmt, ES_FORWARD)
      If bSqlRows Then
         cComCrdLmt = Val(RdoCoCrdLmt!CUWARNCREDITLMT)
      Else
         cComCrdLmt = 0
      End If
      Set RdoCoCrdLmt = Nothing
      
      ' get customer credit limit
      
      sSql = "SELECT CUCREDITLIMIT FROM CustTable" & vbCrLf _
             & "WHERE CUREF ='" & strCust & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCustCL, ES_FORWARD)
      If bSqlRows Then
         With RdoCustCL
            cCusCrdLmt = Val(IIf(IsNull(!CUCREDITLIMIT), 0, !CUCREDITLIMIT))
            ClearResultSet RdoCustCL
         End With
      End If
      
      If (cCusCrdLmt > cComCrdLmt) Then
         MsgBox "warning: Customer Credit Limit exceeds system setting.", vbInformation, MdiSect.ActiveForm.Caption
      End If
   
   
   Set RdoCoCrdLmt = Nothing
   Set RdoCustCL = Nothing
   
   End If
   

End Function
'Public Function AuditSO(strTableName As String, strTransID As String, strITSO As String, strITSONum As String, _
'               strITRev As String, strUser As String, strType As String, strSOColName As String, _
'               strOldValue As String, strNewValue As String)
'
'   If TableExists(strTableName) Then
'      Dim strDate As String
'      strDate = GetServerDate()
'
'      sSql = "INSERT INTO " & strTableName & " (ADT_TRANS_ID, ADT_ITSO, ADT_ITNUMBER, " & _
'               " ADT_ITREV,ADT_USER, ADT_MODIFY_DATE, ADT_MODIFY_TYPE, ADT_SO_COL_NAME, " & _
'               " ADT_OLD_VALUE, ADT_NEW_VALUE) " & _
'               " VALUES ('" & strTransID & "','" & strITSO & "','" & strITSONum & "','" & strITRev & _
'               "','" & strUser & "','" & strDate & "','" & strType & "','" & strSOColName & _
'               "','" & strOldValue & "','" & strNewValue & "')"
'
'      ' Insert the field change
'      InsertAuditEntry strTableName, sSql
'   End If
'
'End Function


Public Function PrintingKanBanLabels(Optional ByVal sCustRef As String) As Boolean
    Dim rdoKanban As ADODB.Recordset
    
    
    PrintingKanBanLabels = False
    ' IF they leave the optional parameter blank, it will return true if any of the customers are setup to print Kanban labels
    If Len(sCustRef) = 0 Then
        sSql = "SELECT CUPRINTKANBAN FROM CustTable WHERE CUPRINTKANBAN=1"
    Else
        sSql = "SELECT CUPRINTKANBAN FROM CustTable WHERE CUREF = '" & sCustRef & "' "
    End If
    Debug.Print sSql
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoKanban, ES_FORWARD)
    
    If bSqlRows Then
        If rdoKanban!CUPRINTKANBAN >= 1 Then PrintingKanBanLabels = True
    End If
    Set rdoKanban = Nothing
End Function

Public Function PrintingPaccarLabels(Optional ByVal sCustRef As String) As Boolean
    Dim rdoPaccar As ADODB.Recordset
    
    
    PrintingPaccarLabels = False
    ' IF they leave the optional parameter blank, it will return true if any of the customers are setup to print Kanban labels
    If Len(sCustRef) = 0 Then
        sSql = "SELECT CUPRINTPACCAR FROM CustTable WHERE CUPRINTPACCAR=1"
    Else
        sSql = "SELECT CUPRINTPACCAR FROM CustTable WHERE CUREF = '" & sCustRef & "' "
    End If
    Debug.Print sSql
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoPaccar, ES_FORWARD)
    
    If bSqlRows Then
        If rdoPaccar!CUPRINTPACCAR >= 1 Then PrintingPaccarLabels = True
    End If
    Set rdoPaccar = Nothing
End Function





Public Sub MakeSureBoxRecordsExist(ByVal sPackSlipNo As String, Optional ByVal sEndPackSlipNo As String)
    Dim RdoBox As ADODB.Recordset
    Dim rdoPS As ADODB.Recordset
    Dim iTotalBoxes As Integer
    Dim sEndingPS As String
    Dim sCurrentPSNumber As String
    Dim i As Integer
    
    If Len(sEndPackSlipNo) = 0 Then sEndingPS = sPackSlipNo Else sEndingPS = sEndPackSlipNo
    
    sSql = "SELECT PSNUMBER, CASE PSBOXES WHEN 0 THEN 1 ELSE PSBOXES END AS TOTBOXES FROM PshdTable " & _
           " WHERE PSNUMBER BETWEEN '" & sPackSlipNo & "' AND '" & sEndingPS & "' "
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_FORWARD)
    
    
   If bSqlRows Then
      With rdoPS
         Do Until .EOF
            iTotalBoxes = !TOTBOXES
            If iTotalBoxes <= 0 Then iTotalBoxes = 1
            sCurrentPSNumber = !PsNumber
            For i = 1 To iTotalBoxes
               sSql = "SELECT * FROM PsibTable WHERE PIBPACKSLIP='" & sCurrentPSNumber & "' AND PIBBOXNO=" & LTrim(str(i))
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoBox, ES_FORWARD)
               If Not bSqlRows Then
                  sSql = "INSERT INTO PsibTable (PIBPACKSLIP, PIBBOXNO, PIBWEIGHT) Values ('" & sCurrentPSNumber & "', " & LTrim(str(i)) & ",0.00)"
                  clsADOCon.ExecuteSql sSql ' rdExecDirect
               End If
               Set RdoBox = Nothing
            Next i
            
            .MoveNext
         Loop
         ClearResultSet rdoPS
      End With
   End If
   
   Set rdoPS = Nothing
    
    'If bSqlRows Then iTotalBoxes = rdoPS!PSBOXES
    'If iTotalBoxes = 0 Then iTotalBoxes = 1
    'Set rdoPS = Nothing
    
    'Dim i As Integer
End Sub
