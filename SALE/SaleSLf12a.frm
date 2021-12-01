VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form SaleSLf12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Sales Order Data from Exostar"
   ClientHeight    =   9690
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   16215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   975
      Left            =   2220
      TabIndex        =   22
      Top             =   2280
      Width           =   3015
      Begin VB.OptionButton optExostar 
         Caption         =   "Exostar New PO"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   200
         Width           =   1815
      End
      Begin VB.OptionButton optExostar 
         Caption         =   "Exostar Change PO"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CheckBox OptSoXml 
      Caption         =   "FromXMLSO"
      Height          =   195
      Left            =   6240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   600
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox cmbPre 
      Height          =   315
      Left            =   2220
      TabIndex        =   12
      Tag             =   "3"
      Text            =   "S"
      ToolTipText     =   "Select or Enter Type A thru Z"
      Top             =   960
      Width           =   520
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2220
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1440
      Width           =   1555
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   2700
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Enter New Sales Order Number"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdSalesOrder 
      Caption         =   "Create/Update SO"
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
      Left            =   14160
      TabIndex        =   9
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   4560
      Width           =   1920
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
      Left            =   14160
      TabIndex        =   8
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   5280
      Width           =   1920
   End
   Begin VB.TextBox txtXMLFilePath 
      Height          =   285
      Left            =   2220
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   3480
      Width           =   6135
   End
   Begin VB.CommandButton cmdImport 
      Cancel          =   -1  'True
      Caption         =   "Import Sales Order Data"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9690
      FormDesignWidth =   16215
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4935
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid GrdChg 
      Height          =   4935
      Left            =   600
      TabIndex        =   25
      Top             =   4680
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   26
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "** Part Not found in Fusion"
      Height          =   255
      Left            =   13920
      TabIndex        =   21
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "* Other Delivery Date"
      Height          =   255
      Left            =   13920
      TabIndex        =   20
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sales Order"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2220
      TabIndex        =   16
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2220
      TabIndex        =   14
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf12a.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf12a.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Exostar File"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   1305
   End
End
Attribute VB_Name = "SaleSLf12a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE

Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim strXML As String
Dim bNewImport As Boolean
Dim ExtName As String

Dim Fields(150) As String

Dim sLastPrefix As String
Dim sNewsonumber As String
Dim sCust As String
Dim sStName As String
Dim sStAdr As String
Dim sContact As String
Dim sConIntPhone As String
Dim sConPhone As String
Dim sConIntFax As String
Dim sConFax As String
Dim sConExt As String
Dim sDivision As String
Dim sOldSoNumber As String
Dim sRegion As String
Dim sSterms As String
Dim sVia As String
Dim sFob As String
Dim sSlsMan As String
Dim sTaxExempt As String
Dim strFilePath As String

Dim iDays As Integer
Dim iFrtDays As Integer
Dim iNetDays As Integer

Dim cDiscount As Currency

'cell lookup by name arrays
Dim CellNames() As String
Dim CellNoCreate() As Integer
Dim cellNoUpdate() As Integer
Dim CellCount As Integer
Dim diagnoseMissingNumCount As Integer

Private txtKeyPress As New EsiKeyBd

Private Sub cmdCan_Click()
   sLastPrefix = cmbPre
   Unload Me

End Sub


Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   cmdImport.Enabled = False
   MouseCursor ccHourglass

   
   On Error GoTo DiaErr1
   strFilePath = txtXMLFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select the XML file to create Sales Order.", _
            vbInformation, Caption
         Exit Sub
   End If

   ExtName = Mid(strFilePath, InStrRev(strFilePath, ".") + 1, Len(strFilePath))

   DeleteOldData ("ExohdImport")
   DeleteOldData ("ExoitImport")
   
   ' process new orders
   GetExcelCellNumbers
   If (optExostar(0).Value = True) Then
      Grd.Visible = True
      GrdChg.Visible = False
      bNewImport = True

      If (ExtName = "xls") Then
        ParseExcelFile (strFilePath)
        FillExcelNewGrid
      Else
        FillGrid (strFilePath)
      End If
      

   ElseIf (optExostar(1).Value = True) Then
      Grd.Visible = False
      GrdChg.Visible = True
      bNewImport = False
      
      If (ExtName = "xls") Then
        ParseExcelFile (strFilePath)
      Else
        ParseChangeOrder (strFilePath)
      End If
      
      FillChgGrid

   Else
      MsgBox "Please select the EDI file type.", _
            vbInformation, Caption
      Exit Sub
   End If
   cmdImport.Enabled = True
   MouseCursor ccArrow
   
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   cmdImport.Enabled = True
   MouseCursor ccArrow
   
End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Text Files (*.txt) | *.txt|" & _
                     "XLS Files (*.xls) | *.xls|" & _
                     "XML Files (*.xml) | *.xml"
   fileDlg.FilterIndex = 2
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtXMLFilePath.Text = ""
   Else
       txtXMLFilePath.Text = fileDlg.FileName
   End If
End Sub

Private Sub cmdSalesOrder_Click()
   
   If Not GetCustomerData(cmbCst.Text) Then
   'If Not bGoodCust Then
      'MsgBox "No customer selected"
      Exit Sub
   End If
   
   If iFrtDays = 0 Then
      If MsgBox("No freight days specified for " & sStName & ".  Proceed anyway?", vbYesNo) <> vbYes Then
         Exit Sub
      End If
   End If
   
   If (optExostar(0).Value = True) Then
     
     Dim bExcel As Boolean
     bExcel = False
     If (ExtName = "xls") Then bExcel = True
    
     CreateNewSOFromXMLData bExcel
      
   ElseIf (optExostar(1).Value = True) Then
      CreateSOFromXMLDataEx
   Else
      MsgBox "Please select the XML file type.", _
            vbInformation, Caption
      Exit Sub
   End If

   Exit Sub
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub
   
Private Function CreateNewSOFromXMLData(bExcel As Boolean)

   Dim bByte As Byte
   Dim lNewSoNum As Long
   Dim iList As Long
   
   Dim strBuyerOrderNumber As String
   Dim strContactName As String
   Dim strContactNum As String
   Dim strContactType As String
   Dim strCusFullName As String
   Dim strShipName As String
   Dim strRefTypeCode As String
   Dim strNewAddress As String
   Dim strPONum As String
   Dim strPartID As String
   Dim strQty As String
   Dim strRefDesc As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strNewSO As String
   Dim strCusName As String
   
   For iList = 1 To Grd.rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 1
         strBuyerOrderNumber = Trim(Grd.Text)
         
         ' Get Customer Ref from Customer Name
         ' NOT needed now
         'GetCustomerRef strCusFullName, strCusName
         strCusName = cmbCst
         ' Get Customer P
         bByte = CheckForCustomerPO(strCusName, strBuyerOrderNumber)
         
         ' if entereing a new item and po exists, inform the user
         If bByte = 1 And optExostar(0).Value = 1 Then
            bByte = MsgBox("The Customer PO Is In Use. Continue?", _
                 ES_YESQUESTION, Caption)
            If bByte = vbNo Then
               Exit Function
            End If
         End If
         
         ' Get new Sales Order number
         Dim strSoType As String
         Dim strItem As String
         strSoType = cmbPre
         
         'GetNewSalesOrder strNewSO, strSoType
         strNewSO = Me.txtSon   '@@@
         
         If (bExcel = True) Then
            CreateSOFromExcelData strFilePath, strBuyerOrderNumber, strNewSO, strSoType, strCusName
         Else
            CreateSOFromXMLData strFilePath, strBuyerOrderNumber, strNewSO, strSoType, strCusName
         End If
         
         If optSORev.Value = vbChecked Then
            OptSoXml = vbChecked
            SaleSLe02a.Show
            SaleSLe02a.OptSoXml = vbChecked
            SaleSLe02a.SetFocus
            SaleSLe02a.cmbSon.SetFocus
         End If
      
      End If
   Next
   
End Function


'Private Sub AddSalesOrder(strNewSO As String, strBuyerOrderNumber As String, _
'                     strContactName As String, strContactNum As String _
'                     , strShipName As String, strCusName As String, strNewAddress As String, _
'                     strSoType As String, strSORemark As String)
'
'   Dim sNewDate As Variant
'   Dim bGoodCust As Byte
'
'   If Len(strContactName) > 20 Then
'      strContactName = Mid(strContactName, 1, 20)
'   End If
'
'   bGoodCust = GetCustomerData(strCusName)
'   If bCutOff = 1 Then
'      MsgBox "This Customer's Credit Is On Hold.", _
'         vbInformation, Caption
'      bGoodCust = 0
'   End If
'   If Not bGoodCust Then Exit Sub
'   On Error GoTo DiaErr1
'
'   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
'   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
'          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
'          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
'          & "SOTEXT,SOTAXEXEMPT,SOPO, SOREMARKS) " _
'          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
'          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
'          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
'          & sRegion & "','" & sSterms & "','" & sVia & "','" _
'          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
'          & "," & iFrtDays & ",'" & strNewSO & "','" & sTaxExempt & "','" _
'          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
'
'   'Debug.Print sSql
'
'   clsADOCon.ExecuteSql sSql ' rdExecDirect
'   If clsADOCon.RowsAffected Then
'      On Error Resume Next
'      MsgBox "Sales Order Added - " & strNewSO, vbInformation, Caption
'      sSql = "UPDATE SohdTable SET SOCCONTACT='" & strContactName & "'," _
'             & "SOCPHONE='" & strContactNum & "',SOCINTFAX='" & sConIntFax _
'             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
'             & " WHERE SONUMBER=" & Val(strNewSO) & ""
'      'Debug.Print sSql
'
'      clsADOCon.ExecuteSql sSql ' rdExecDirect
'
'      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(strSoType) _
'             & Trim(strNewSO) & "' WHERE COREF=1"
'      clsADOCon.ExecuteSql sSql ' rdExecDirect
'
'      ''@@@ update sales order info
'      lblLst = Me.cmbPre + strNewSO
'      txtSon = CStr(CLng(strNewSO) + 1)
'
'   Else
'      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
'   End If
'   Exit Sub
'
'DiaErr1:
'   MsgBox Err.Description
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub

Private Sub AddSalesOrder(strNewSO As String, strBuyerOrderNumber As String, _
                     strContactName As String, strContactNum As String _
                     , strShipName As String, strCusName As String, strNewAddress As String, _
                     strSoType As String, strSORemark As String)
                     
   Dim sNewDate As Variant
   Dim bGoodCust As Byte
   
   If Len(strContactName) > 20 Then
      strContactName = Mid(strContactName, 1, 20)
   End If
   
   bGoodCust = GetCustomerData(strCusName)
   If bCutOff = 1 Then
      MsgBox "This Customer's Credit Is On Hold.", _
         vbInformation, Caption
      bGoodCust = 0
   End If
   If Not bGoodCust Then Exit Sub
   On Error GoTo DiaErr1
   
   'make sure no one is simultaneously creating the same SO number
   Dim so As Long
   so = CLng(strNewSO)
   
   clsADOCon.BeginTrans
   Dim rs As ADODB.Recordset
   Do While True
      sSql = "select SONUMBER from SohdTable where SONUMBER = " & so
      If clsADOCon.GetDataSet(sSql, rs) = 0 Then Exit Do
      so = so + 1
      strNewSO = so
   Loop
   
   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
'   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
'          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
'          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
'          & "SOTEXT,SOTAXEXEMPT,SOPO, SOREMARKS) " _
'          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
'          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
'          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
'          & sRegion & "','" & sSterms & "','" & sVia & "','" _
'          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
'          & "," & iFrtDays & ",'" & strNewSO & "','" & sTaxExempt & "','" _
'          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
'
   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
          & "SOTAXEXEMPT,SOPO, SOREMARKS) " _
          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
          & sRegion & "','" & sSterms & "','" & sVia & "','" _
          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
          & "," & iFrtDays & ",'" & sTaxExempt & "','" _
          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   If clsADOCon.RowsAffected Then
      On Error Resume Next
      MsgBox "Sales Order Added - " & strNewSO, vbInformation, Caption
      sSql = "UPDATE SohdTable SET SOCCONTACT='" & strContactName & "'," _
             & "SOCPHONE='" & strContactNum & "',SOCINTFAX='" & sConIntFax _
             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
             & " WHERE SONUMBER=" & Val(strNewSO) & ""
      'Debug.Print sSql
      
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      
      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(strSoType) _
             & Trim(strNewSO) & "' WHERE COREF=1"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      clsADOCon.CommitTrans
      
      ''@@@ update sales order info
      lblLst = Me.cmbPre + strNewSO
      txtSon = CStr(CLng(strNewSO) + 1)
   
   Else
      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   MsgBox Err.Description
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub AddSoItem(strNewSO As String, strItem As String, _
                     strBuyerOrderNumber As String, strBuyerItemLine As String, _
                     strPartID As String, strQty As String, _
                     strUnitPrice As String, strReqDt As String)
   Dim NewReqDt As Date
   Dim strNewReqDt As String
   
   NewReqDt = DateAdd("d", -iFrtDays, CDate(strReqDt))
   
   strNewReqDt = Format(NewReqDt, "mm/dd/yyyy")
   
   ' MM clsADOCon.BeginTrans
   ' MM clsADOCon.ADOErrNum = 0
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITCUSTITEMNO, ITPART,ITQTY,ITSCHED,ITBOOKDATE," _
          & "ITCUSTREQ,ITSCHEDDEL, ITDOLLORIG, ITDOLLARS, ITUSER) " _
          & "VALUES(" & strNewSO & "," & strItem & ",'" & strBuyerItemLine & "','" _
          & Compress(strPartID) & "'," & Val(strQty) & ",'" & strNewReqDt & "','" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & strReqDt & "','" _
          & strReqDt & "','" & CCur(strUnitPrice) & "','" _
          & CCur(strUnitPrice) & "','" & sInitials & "')"
          
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   'Add commission if applicable.
'   If cmdCom.Enabled Then
     Dim Item As New ClassSoItem
     Dim bUserMsg As Boolean
     bUserMsg = False
     Item.InsertCommission CLng(strNewSO), CLng(strItem), "", ""
     Item.UpdateCommissions CLng(strNewSO), CLng(strItem), "", bUserMsg
 '  End If

   ' MM clsADOCon.CommitTrans
   
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPartComm(ByVal strGetPart As String, _
            ByRef strPartNum As String, ByRef bComm As Boolean) As Byte
   Dim RdoPrt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   bComm = False
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            strPartNum = "" & Trim(!PartNum)
            If !PACOMMISSION = 1 Then bComm = True _
                               Else bComm = False
            GetPartComm = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPartComm = 0
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      GetPartComm = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPartComm"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetCustomerData(strCusName As String) As Byte
   Dim RdoCst As ADODB.Recordset
   sCust = Compress(strCusName)
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & strCusName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & " " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP)
         sDivision = "" & Trim(!CUDIVISION)
         sRegion = "" & Trim(!CUREGION)
         sSterms = "" & Trim(!CUSTERMS)
         sVia = "" & Trim(!CUVIA)
         sFob = "" & Trim(!CUFOB)
         sSlsMan = "" & Trim(!CUSALESMAN)
         sContact = "" & Trim(!CUCCONTACT)
         sConIntPhone = "" & Trim(!CUCINTPHONE)
         sConPhone = "" & Trim(!CUCPHONE)
         sConIntFax = "" & Trim(!CUINTFAX)
         sConFax = "" & Trim(!CUFAX)
         sConExt = "" & Trim(str$(!CUCEXT))
         cDiscount = Format(0 + !CUARDISC, "##0.000")
         iDays = Format(!CUDAYS, "###0")
         iNetDays = Format(!CUNETDAYS, "###0")
         iFrtDays = Format(!CUFRTDAYS, "##0")
         sTaxExempt = "" & Trim(!CUTAXEXEMPT)
         ClearResultSet RdoCst
      End With
      GetCustomerData = True
   Else
      sStName = ""
      sStAdr = ""
      sDivision = ""
      sRegion = ""
      sSterms = ""
      sVia = ""
      sFob = ""
      sSlsMan = ""
      iFrtDays = 0
      MsgBox "Couldn't Retrieve Customer.", vbExclamation, Caption
      GetCustomerData = False
   End If
   Set RdoCst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcustda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetCustomerRef(ByRef strCusFullName As String, ByRef strCusName As String)

   Dim RdoCus As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF FROM CustTable WHERE CUNAME = '" & strCusFullName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCus)
   If bSqlRows Then
      With RdoCus
         strCusName = Trim(!CUREF)
         ClearResultSet RdoCus
      End With
   Else
      strCusName = ""
   End If
   Set RdoCus = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetCustomerRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'GetCustomerRef = False
   DoModuleErrors Me
   

End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst, False
   lblNotice.Visible = False
End Sub


Private Sub cmbPre_LostFocus()
   Dim a As Integer
   cmbPre = CheckLen(cmbPre, 1)
   On Error Resume Next
   a = Asc(Left(cmbPre, 1))
   If a < 65 Or a > 90 Then
      MsgBox "Must Be Between A and Z..", vbInformation, Caption
      cmbPre = sLastPrefix
   End If
   If Len(Trim(cmbPre)) = 0 Then cmbPre = sLastPrefix
   
End Sub


Private Sub Form_Activate()
Dim bSoAdded As Byte
   MdiSect.lblBotPanel = Caption
   
   GetLastSalesOrder sOldSoNumber, sNewsonumber, True
   FillCustomers
   'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
   cmbCst.Text = GetSetting("Esi2000", "EsiSale", "Exostar_Cust", "")

   FindCustomer Me, cmbCst, False
   bSoAdded = 0
   OptSoXml.Value = vbUnchecked
   'tmr1.Enabled = True
  
   If bOnLoad Then
       bOnLoad = 0
   End If
    
    
   MouseCursor (0)

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   Call WheelUnHook(Me.hWnd)
   SaveSetting "Esi2000", "EsiSale", "LastPrefix", sLastPrefix
   If cmbCst.Text <> "" Then
      SaveSetting "Esi2000", "EsiSale", "Exostar_Cust", cmbCst.Text
   End If
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
    CellCount = 0
   
   Dim iChar As Integer
   sLastPrefix = GetSetting("Esi2000", "EsiSale", "LastPrefix", sLastPrefix)
   If Len(sLastPrefix) = 0 Then sLastPrefix = "S"
   cmbPre = sLastPrefix
   For iChar = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(iChar)
   Next
    
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
   
      .rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "PO Number"
      .Col = 2
      .Text = "Line Item"
      .Col = 3
      .Text = "PartNumber"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "UnitPrice"
      .Col = 6
      .Text = "Requestdate"
      .Col = 7
      .Text = "ShipTo Name"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1250
      .ColWidth(2) = 500
      .ColWidth(3) = 2500
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 3250
'      .ColWidth(8) = 1000
'      .ColWidth(9) = 1000
'      .ColWidth(10) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   With GrdChg
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
   
      .rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "SO Number"
      .Col = 2
      .Text = "PO Number"
      .Col = 3
      .Text = "Line Item"
      .Col = 4
      .Text = "ChangeData"
      .Col = 5
      .Text = "PartNumber"
      .Col = 6
      .Text = "Qty"
      .Col = 7
      .Text = "UnitPrice"
      .Col = 8
      .Text = "Requestdate"
      .Col = 9
      .Text = "ShipTo Name"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1050
      .ColWidth(2) = 1250
      .ColWidth(3) = 1050
      .ColWidth(4) = 500
      .ColWidth(5) = 2500
      .ColWidth(6) = 1000
      .ColWidth(7) = 1200
      .ColWidth(8) = 1200
      .ColWidth(9) = 3250
'      .ColWidth(8) = 1000
'      .ColWidth(9) = 1000
'      .ColWidth(10) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Function FillGrid(ByVal strFilePath As String) As Integer
   
   MouseCursor ccHourglass
   Grd.rows = 1
   On Error GoTo DiaErr1
       
   ' Read the content if the text file.
   Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
   Dim lngPos As Integer
   Dim bFound As Boolean
' Get a free file number
   nFileNum = FreeFile
   
   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   bFound = False
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sNextLine
      sNextLine = sNextLine & vbCrLf
      sText = sText & sNextLine
   Loop
   Close nFileNum

'   Do While Not EOF(nFileNum)
'      Line Input #nFileNum, sNextLine
'      lngPos = InStr(sNextLine, "<Order>")
'      If bFound = False And lngPos = 0 Then
'         sText = ""
'         sNextLine = ""
'      Else
'         sNextLine = sNextLine & vbCrLf
'         sText = sText & sNextLine
'         bFound = True
'      End If
'
'   Loop
   'Input #1, strXML
   'Close nFileNum
   ' Add the Orderlist collections
   strXML = "<OrderList>" & sText & "</OrderList>"
   
   Dim OrderDoc As MSXML2.DOMDocument40  'XML document object
       
   'Create a new document object
   Set OrderDoc = New MSXML2.DOMDocument40
   OrderDoc.preserveWhiteSpace = False
   OrderDoc.async = False
   OrderDoc.validateOnParse = False
   
   OrderDoc.resolveExternals = True
   'Remove the cached schema, we'll restore it later if needed
   Set OrderDoc.schemas = Nothing
   
       
   OrderDoc.loadXML (strXML)
   If OrderDoc.parseError.errorCode <> 0 Then
      MsgBox "Document could not be parsed:" & vbCrLf & OrderDoc.parseError.reason
      Exit Function
   End If
        
      On Error Resume Next
      'Update the book list shown in the Treeview to match any changes
      Dim OrderNode As MSXML2.IXMLDOMNode      'the "book" node
      
      Dim strContactName, strContactNum, strContactType, strShipName1, strShipName2 As String
      Dim strStreet, strCustName, strStreetSup1, strStreetSup2, strPostalCode, strCity, strRegionCode As String
      Dim strPartID, strQty, strUOMQty, strRefTypeCode, strRefNum As String
      Dim strRefDesc, strUnitPrice As String
      Dim strRequestDlvDate As String
      Dim strReqDt As String
      Dim strYear, strMonth, strDay As String
      Dim bDeliveryDate As Boolean
      Dim bPartFound As Boolean
      Dim strTotQty As String
      Dim strSchedID As String
      Dim strBuyerPartyAddressName As String
      Dim strShpToLocName1 As String
      Dim strShpToPtyName1 As String
      Dim strIdent As String
      
      Dim i As Integer
      Dim strBuyerOrderNumber As String
      Dim bIncRow As Boolean
      Dim iItem  As Integer
      For Each OrderNode In OrderDoc.documentElement.childNodes 'OrderDoc.selectNodes("//OrderList/Order")
          
         If OrderNode.baseName = "Order" Then
            strBuyerOrderNumber = OrderNode.selectSingleNode("./OrderHeader/OrderNumber/BuyerOrderNumber").Text
            
            Grd.rows = Grd.rows + 1
            Grd.Row = Grd.rows - 1
            bIncRow = False
            iItem = 1
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = Trim(strBuyerOrderNumber)
            
            Dim ItemDetailsNode As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
            Dim ItemDetailNode As MSXML2.IXMLDOMNode
            
            Set ItemDetailsNode = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail")
            
            For Each ItemDetailNode In ItemDetailsNode.childNodes
            
               If bIncRow = True Then
                  Grd.rows = Grd.rows + 1
                  Grd.Row = Grd.rows - 1
                  iItem = iItem + 1
                  bIncRow = False
               End If
               
               strRequestDlvDate = ""
               strPartID = ""
               strUnitPrice = ""
               strQty = ""
               bDeliveryDate = False
               strShpToLocName1 = ""
               strShpToPtyName1 = ""
               
               strBuyerPartyAddressName = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/NameAddress/Name1").Text
               strShpToLocName1 = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Name1").Text
               strShpToPtyName1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name1").Text
               
               'strShipToNameAdr1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name1").Text
               'strShipToNameAdr2 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name2").Text
               
               strPartID = ItemDetailNode.selectSingleNode("./BaseItemDetail/ItemIdentifiers/PartNumbers/BuyerPartNumber/PartNum/PartID").Text
               strTotQty = ItemDetailNode.selectSingleNode("./BaseItemDetail/TotalQuantity/Quantity/QuantityValue").Text
               strUnitPrice = ItemDetailNode.selectSingleNode("./PricingDetail/ListOfPrice/Price/UnitPrice/UnitPriceValue").Text
               strRequestDlvDate = ItemDetailNode.selectSingleNode("./DeliveryDetail/ListOfScheduleLine/ScheduleLine/RequestedDeliveryDate").Text
               strIdent = ItemDetailNode.selectSingleNode("./BaseItemDetail/ItemIdentifiers/CommodityCode/Identifier/Ident").Text
            
               Dim nodeListSchLine As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
               Dim nodeScheduleLine As MSXML2.IXMLDOMNode
               
               Set nodeListSchLine = ItemDetailNode.selectSingleNode("./DeliveryDetail/ListOfScheduleLine")
               
               For Each nodeListSchLine In nodeListSchLine.childNodes
                  
                  If bIncRow = True Then
                     Grd.rows = Grd.rows + 1
                     Grd.Row = Grd.rows - 1
                     iItem = iItem + 1
                     bIncRow = False
                  End If
                  
                  strQty = nodeListSchLine.selectSingleNode("./Quantity/QuantityValue").Text
                  strSchedID = nodeListSchLine.selectSingleNode("./ScheduleLineID").Text
            
                  Grd.Col = 2
                  Grd.Text = Trim(CStr(strIdent))
                  
                  sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartID) & "'"
                  bPartFound = CheckRecordExits(sSql)
                  Grd.Col = 3
                  If (bPartFound = False) Then
                     Grd.Text = "**" & Trim(strPartID)
                  Else
                     Grd.Text = Trim(strPartID)
                  End If
                  
                  Grd.Col = 4
                  Grd.Text = Trim(strQty)
                  Grd.Col = 5
                  Grd.Text = Trim(strUnitPrice)
                  Grd.Col = 6
                  
                  GetDeliveryDate nodeListSchLine, strRequestDlvDate, strReqDt, bDeliveryDate
                  If (bDeliveryDate = False) Then
                     Grd.Text = "*" & Trim(strReqDt)
                  Else
                     Grd.Text = Trim(strReqDt)
                  End If
                  
                  Dim strShipAddr As String
                  
                  If (strShpToLocName1 <> "") Then
                     strShipAddr = strShpToLocName1
                  Else
                     strShipAddr = strShpToPtyName1
                  End If
                  
                  Grd.Col = 7
                  Grd.Text = Trim(strShipAddr)
                  
                  bIncRow = True
               Next
               
               bIncRow = True
            Next ' Item Detail
         
            Grd.rows = Grd.rows + 1
            Grd.Row = Grd.rows - 1
         End If   ' if not order
      Next ' Order
   MouseCursor ccArrow
   
   Set OrderDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function FillExcelNewGrid() As Integer
   
    MouseCursor ccHourglass
    Grd.rows = 1
    On Error GoTo DiaErr1
     
'MsgBox ("Inside FillExcelNewGrid:")
      
    On Error Resume Next
    'Update the book list shown in the Treeview to match any changes
    
    Dim strContactName, strContactNum, strContactType, strShipName1, strShipName2 As String
    Dim strStreet, strCustName, strStreetSup1, strStreetSup2, strPostalCode, strCity, strRegionCode As String
    Dim strPartID, strQty, strUOMQty, strRefTypeCode, strRefNum As String
    Dim strRefDesc, strUnitPrice As String
    Dim strRequestDlvDate As String
    Dim strReqDt As String
    Dim strYear, strMonth, strDay As String
    Dim bDeliveryDate As Boolean
    Dim bPartFound As Boolean
    Dim strTotQty As String
    Dim strSchedID As String
    Dim strBuyerPartyAddressName As String
    Dim strShpToLocName1 As String
    Dim strShpToPtyName1 As String
    Dim strIdent As String
    
    Dim i As Integer
    Dim strBuyerOrderNumber As String
    Dim bIncRow As Boolean
    Dim iItem  As Integer
      
      
    sSql = "SELECT DISTINCT a.SOPO_BUYERORDNUM, a.EXOSTART_IMPORT_TYPE, " _
           & "ISNULL(INDEX_NUM, 0) as INDEX_NUM, BUYER_PARTY_ADDR_NAME1, " _
           & "SHPTO_PARTY_ADDRNAME1, a.ORDDET_SHPTO_LOC_ADDRNAME1," _
           & " convert(int, BUYER_LINEITEM_NUM) as buyerlineNum, PART_ID, TOT_QTY," _
           & "UNIT_PRICE, REQ_DELDATE,SCHED_QTY_VALUE, SCHED_LINE_ID " _
        & " FROM ExohdImport a, ExoitImport b " _
           & " WHERE a.SOPO_BUYERORDNUM = b.SOPO_BUYERORDNUM " _
           & " AND a.EXOSTART_IMPORT_TYPE = b.EXOSTART_IMPORT_TYPE" _
           & " ORDER BY a.SOPO_BUYERORDNUM, buyerlineNum, a.EXOSTART_IMPORT_TYPE"
    
    'Debug.Print sSql

'MsgBox ("Exostar SQL:" + sSql)
    
    Dim RdoExo As ADODB.Recordset
    Dim strPrevBuyOrdNum As String
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
      
      
   Grd.rows = 1
   If bSqlRows Then
      With RdoExo
      While Not .EOF
        strBuyerOrderNumber = !SOPO_BUYERORDNUM
        
'MsgBox ("Inside Dataset:" + strBuyerOrderNumber)
        
        Grd.rows = Grd.rows + 1
        Grd.Row = Grd.rows - 1
        bIncRow = False
        iItem = 1
        
         If (strBuyerOrderNumber <> strPrevBuyOrdNum) Then
            
            GrdChg.Col = 0
            Set GrdChg.CellPicture = Chkno.Picture
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = Trim(strBuyerOrderNumber)
         
            strPrevBuyOrdNum = strBuyerOrderNumber
         End If
        
'        Grd.Col = 0
'        Set Grd.CellPicture = Chkno.Picture
'        Grd.Col = 1
'        Grd.Text = Trim(strBuyerOrderNumber)
        
        strRequestDlvDate = ""
        strPartID = ""
        strUnitPrice = ""
        strQty = ""
        bDeliveryDate = False
        strShpToLocName1 = ""
        strShpToPtyName1 = ""
               
        strBuyerPartyAddressName = !BUYER_PARTY_ADDR_NAME1
        strShpToLocName1 = !ORDDET_SHPTO_LOC_ADDRNAME1
        strShpToPtyName1 = !SHPTO_PARTY_ADDRNAME1
        
        strPartID = !PART_ID
        strTotQty = !TOT_QTY
        strUnitPrice = !UNIT_PRICE
        strRequestDlvDate = !REQ_DELDATE
        strIdent = ""
            
        strQty = !SCHED_QTY_VALUE
        strSchedID = !SCHED_LINE_ID
        
        Grd.Col = 2
        Grd.Text = Trim(CStr(!buyerlineNum))
        
        sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartID) & "'"
        bPartFound = CheckRecordExits(sSql)
        Grd.Col = 3
        If (bPartFound = False) Then
        Grd.Text = "**" & Trim(strPartID)
        Else
        Grd.Text = Trim(strPartID)
        End If
        
        Grd.Col = 4
        Grd.Text = Trim(strQty)
        Grd.Col = 5
        Grd.Text = Trim(strUnitPrice)
        Grd.Col = 6
        Grd.Text = Trim(strRequestDlvDate)
        
        Dim strShipAddr As String
        
        If (strShpToLocName1 <> "") Then
            strShipAddr = strShpToLocName1
        Else
            strShipAddr = strShpToPtyName1
        End If
        
        Grd.Col = 7
        Grd.Text = Trim(strShipAddr)
        
        bIncRow = True
        'Grd.Rows = Grd.Rows + 1
        'Grd.Row = Grd.Rows - 1
        
        .MoveNext
      Wend
      .Close
      End With
   End If
   MouseCursor ccArrow
   
'MsgBox ("Done FillExcelNewGrid:")
   
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function FillChgGrid() As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   'Update the book list shown in the Treeview to match any changes
   Dim OrderNode As MSXML2.IXMLDOMNode      'the "book" node
   
   Dim strContactName, strContactNum, strContactType, strShipName1, strShipName2 As String
   Dim strStreet, strCustName, strStreetSup1, strStreetSup2, strPostalCode, strCity, strRegionCode As String
   Dim strPartID, strQty, strUOMQty, strRefTypeCode, strRefNum As String
   Dim strRefDesc, strUnitPrice As String
   Dim strRequestDlvDate As String
   Dim bDeliveryDate As Boolean
   Dim bPartFound As Boolean
   Dim strTotQty As String
   Dim strSchedID As String
   Dim strBuyerPartyAddressName As String
   Dim strShpToLocName1 As String
   Dim strShpToPtyName1 As String
   Dim strImpType As String
   Dim strBuyItmNum As String
   Dim strPrevBuyOrdNum As String
   Dim strPrevBuyItmNum As String
   
   Dim strSONumber As String
   
   Dim i As Integer
   Dim strBuyerOrderNumber As String
   Dim bIncRow As Boolean
   Dim iItem  As Integer
   
   sSql = "SELECT DISTINCT SONUMBER, a.SOPO_BUYERORDNUM, a.EXOSTART_IMPORT_TYPE," & vbCrLf _
      & "ISNULL(INDEX_NUM, 0) as INDEX_NUM, BUYER_PARTY_ADDR_NAME1," & vbCrLf _
      & "SHPTO_PARTY_ADDRNAME1, a.ORDDET_SHPTO_LOC_ADDRNAME1," & vbCrLf _
      & "BUYER_LINEITEM_NUM , PART_ID, TOT_QTY," & vbCrLf _
      & "cast(cast(UNIT_PRICE as decimal(12,2)) as varchar(10)) as UNIT_PRICE," & vbCrLf _
      & "convert(varchar(10),cast(REQ_DELDATE as date),101) as REQ_DELDATE," & vbCrLf _
      & "SCHED_QTY_VALUE, SCHED_LINE_ID," & vbCrLf _
      & "Convert(int, BUYER_LINEITEM_NUM) as buyerlineitem," & vbCrLf _
      & "isnull(CONVERT(varchar(10),ITCUSTREQ,101),'') as ITCUSTREQ," & vbCrLf _
      & "isnull(cast(cast(ITQTY as int) as varchar(10)),'') as ITQTY," & vbCrLf _
      & "isnull(cast(cast(ITDOLLARS as decimal(12,2)) as varchar(10)),'') as ITDOLLARS," & vbCrLf _
      & "isnull(cast(ITNUMBER as varchar(5)),'') as ITNUMBER" & vbCrLf _
      & "FROM ExohdImport a" & vbCrLf _
      & "join ExoitImport b on a.SOPO_BUYERORDNUM = b.SOPO_BUYERORDNUM" & vbCrLf _
      & "and a.EXOSTART_IMPORT_TYPE = b.EXOSTART_IMPORT_TYPE" & vbCrLf _
      & "join SohdTable so on so.SOPO = a.SOPO_BUYERORDNUM" & vbCrLf _
      & "left join SoitTable si on si.ITSO = so.SONUMBER" & vbCrLf _
      & "and si.ITCUSTITEMNO = b.BUYER_LINEITEM_NUM" & vbCrLf _
      & "ORDER BY a.SOPO_BUYERORDNUM, buyerlineitem, a.EXOSTART_IMPORT_TYPE"

   'Debug.Print sSql
   
   Dim RdoExo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
   
   GrdChg.rows = 1
   If bSqlRows Then
      With RdoExo
      While Not .EOF
         
         strSONumber = !SoNumber
         strBuyerOrderNumber = !SOPO_BUYERORDNUM
         strImpType = !EXOSTART_IMPORT_TYPE
         strBuyerPartyAddressName = !BUYER_PARTY_ADDR_NAME1
         strShpToPtyName1 = !SHPTO_PARTY_ADDRNAME1
         strShpToLocName1 = !ORDDET_SHPTO_LOC_ADDRNAME1
         iItem = CInt(!INDEX_NUM)
         strBuyItmNum = !BUYER_LINEITEM_NUM
         strPartID = !PART_ID
         strTotQty = !TOT_QTY
         strUnitPrice = !UNIT_PRICE
         strRequestDlvDate = !REQ_DELDATE
         
         
         strQty = !SCHED_QTY_VALUE
         strSchedID = !SCHED_LINE_ID
               
         GrdChg.rows = GrdChg.rows + 1
         GrdChg.Row = GrdChg.rows - 1
         
         If ((strBuyerOrderNumber <> strPrevBuyOrdNum) Or _
            (strBuyItmNum <> strPrevBuyItmNum)) Then
            
            GrdChg.Col = 0
            Set GrdChg.CellPicture = Chkno.Picture
            
            GrdChg.Col = 1
            GrdChg.Text = Trim(strSONumber)
            GrdChg.Col = 2
            GrdChg.Text = Trim(strBuyerOrderNumber)
            GrdChg.Col = 3
            GrdChg.Text = Trim(strBuyItmNum)
         
            strPrevBuyOrdNum = strBuyerOrderNumber
            strPrevBuyItmNum = strBuyItmNum
            
         End If
         
         GrdChg.Col = 4
         GrdChg.Text = Trim(strImpType)
         
         sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartID) & "'"
         bPartFound = CheckRecordExits(sSql)
         
         GrdChg.Col = 5
         If (bPartFound = False) Then
            GrdChg.Text = "**" & Trim(strPartID)
         Else
            GrdChg.Text = Trim(strPartID)
         End If
         
         GrdChg.Col = 6
         GrdChg.Text = Trim(strQty)
         GrdChg.Col = 7
         GrdChg.Text = Trim(strUnitPrice)
         GrdChg.Col = 8
         GrdChg.Text = Trim(strRequestDlvDate)
         
         Dim strShipAddr As String
         
         If (strShpToLocName1 <> "") Then
            strShipAddr = strShpToLocName1
         Else
            strShipAddr = strShpToPtyName1
         End If
         
         GrdChg.Col = 9
         GrdChg.Text = Trim(strShipAddr)
         
         'add original parameters with different row color
         GrdChg.rows = GrdChg.rows + 1
         GrdChg.Row = GrdChg.rows - 1
         Dim intcols As Integer
         For intcols = 0 To GrdChg.Cols - 1
             GrdChg.Col = intcols
             GrdChg.CellBackColor = &HF0F0F0
         Next intcols
         
         GrdChg.Col = 3
         GrdChg.Text = !ITNUMBER
         GrdChg.Col = 6
         GrdChg.Text = !ITQty
         GrdChg.Col = 7
         GrdChg.Text = !ITDOLLARS
         GrdChg.Col = 8
         GrdChg.Text = !itcustreq
                 
         .MoveNext
      Wend
      .Close
      End With
   End If
   
   Set RdoExo = Nothing
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "fillchggrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function ParseChangeOrder(ByVal strFilePath As String) As Integer
   
   MouseCursor ccHourglass
   Grd.rows = 1
   On Error GoTo DiaErr1
       
   ' Read the content if the text file.
   Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
   Dim lngPos As Integer
   Dim bFound As Boolean
' Get a free file number
   nFileNum = FreeFile

   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   bFound = False
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sNextLine
      sNextLine = sNextLine & vbCrLf
      sText = sText & sNextLine
   Loop
   Close nFileNum

   
   strXML = "<ChangeOrderList>" & sText & "</ChangeOrderList>"
   Dim OrderDoc As MSXML2.DOMDocument40  'XML document object
       
   'Create a new document object
   Set OrderDoc = New MSXML2.DOMDocument40
   OrderDoc.preserveWhiteSpace = False
   OrderDoc.async = False
   OrderDoc.validateOnParse = False
   
   OrderDoc.resolveExternals = True
   'Remove the cached schema, we'll restore it later if needed
   Set OrderDoc.schemas = Nothing
   
       
   OrderDoc.loadXML (strXML)
   If OrderDoc.parseError.errorCode <> 0 Then
      MsgBox "Document could not be parsed:" & vbCrLf & OrderDoc.parseError.reason
      Exit Function
   End If
        
   On Error Resume Next
   'Update the book list shown in the Treeview to match any changes
   Dim ChgOrderNode As MSXML2.IXMLDOMNode      'the "book" node
   
   Dim strContactName, strContactNum, strContactType, strShipName1, strShipName2 As String
   Dim strStreet, strCustName, strStreetSup1, strStreetSup2, strPostalCode, strCity, strRegionCode As String
   Dim strPartID, strQty, strUOMQty, strRefTypeCode, strRefNum As String
   Dim strRefDesc, strUnitPrice As String
   Dim strRequestDlvDate As String
   Dim strReqDt As String
   Dim strYear, strMonth, strDay As String
   Dim bDeliveryDate As Boolean
   Dim bPartFound As Boolean
   
   Dim i As Integer
   Dim bIncRow As Boolean
   Dim iItem  As Integer
   Dim strPreOrgPath As String
   Dim strPreChgPath  As String
   Dim strBuyerOrderNum As String
   
   For Each ChgOrderNode In OrderDoc.documentElement.childNodes 'OrderDoc.selectNodes("//OrderList/Order")
      
      If ChgOrderNode.baseName = "ChangeOrder" Then
         
         ' Get the Buyer Number (PO)
         strPreOrgPath = "./ChangeOrderHeader/OriginalOrderHeader/"
         GetNodeText ChgOrderNode, strBuyerOrderNum, GetNodeElemPath("SOPO_BUYERORDNUM", strPreOrgPath)
         
         ImportHeaderData ChgOrderNode, "ORG", strPreOrgPath
         
         strPreChgPath = "./ChangeOrderHeader/OrderHeaderChanges/"
         ImportHeaderData ChgOrderNode, "CHG", strPreChgPath
         
         Dim ChgItemDetailsNode As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
         Dim ChgItemDetailNode As MSXML2.IXMLDOMNode
         
         Set ChgItemDetailsNode = ChgOrderNode.selectSingleNode("./ChangeOrderDetail/ListOfChangeOrderItemDetail")
         
         Dim Index As Integer
         Index = 1
         For Each ChgItemDetailNode In ChgItemDetailsNode.childNodes
            
            
            Dim OrgItemDetail As MSXML2.IXMLDOMNode
            Set OrgItemDetail = ChgItemDetailNode.selectSingleNode(GetNodeElemPath("CHG_ORG_ITEMDETAIL_NODE"))
            If (Not OrgItemDetail Is Nothing) Then
               ' Original Order detail
               GetItemDetailNode ChgOrderNode, OrgItemDetail, strBuyerOrderNum, "ORG", Index
            End If
            
            Dim ItemDetailChg As MSXML2.IXMLDOMNode
            Set ItemDetailChg = ChgItemDetailNode.selectSingleNode(GetNodeElemPath("CHG_ITEMDETAIL_NODE"))
            
            If (Not ItemDetailChg Is Nothing) Then
               ' Change Order detail
               GetItemDetailNode ChgOrderNode, ItemDetailChg, strBuyerOrderNum, "CHG", Index
            End If
            
            Index = Index + 1
            
         Next ' Item Detail
      
         Grd.rows = Grd.rows + 1
         Grd.Row = Grd.rows - 1
      End If   ' if not order
   Next ' Order
   
   OrderDoc.abort
   Set OrderDoc = Nothing
   
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "ParseChangeOrder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function ParseExcelFile(ByVal strFullPath As String) As Integer
   'StopwatchStart
    Dim xlApp As Excel.Application
   'Dim xlApp As Object
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim strPartNum As String
   Dim strPartType As String
   Dim strPartCode As String
   Dim strPartClass As String
   Dim iIndex As Integer
   Dim bContinue As Boolean
   Dim strOrderType As String
   Dim strBuyerOrderNum As String
   
   On Error GoTo DiaErr1
   
'MsgBox ("Full Path :" + strFullPath)
   If (strFullPath <> "") Then
      'Set xlApp = New Excel.Application
      Set xlApp = CreateObject("Excel.Application")
'MsgBox ("Excel Application:" + xlApp.Name)
   
      Set wb = xlApp.Workbooks.open(strFullPath)
'MsgBox ("Wb Name :" + wb.Name)
   
      Set ws = wb.Worksheets(1) 'Specify your worksheet name
'MsgBox ("Work sheet :" + ws.Name)
      
      bContinue = True
      iIndex = 2
      
      While (bContinue)
         
        ReadAllFields iIndex, ws

        If (Fields(0) <> "") Then
            strOrderType = GetExcelCellValue("ORDER_TYPE")
            
            If (bNewImport = True) Then
    
                If (strOrderType <> "Updated") Then
                    
                    strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
                    
                    ImportExcelHeaderData Fields, "NewOrder"
                    
                    GetExcelItemDetail strBuyerOrderNum, "NewOrder"
                End If
            Else
            
                If (strOrderType = "Updated") Then
                     ' 4/5/2020: Only schedule status = "Updated" also
                     ' Buyer Company Region = "Boeing Commercial Airplanes"
                     ' Buyer Company = "The Boeing Company"
                     Dim scheduleStatus As String, buyerCo As String
                     scheduleStatus = GetExcelCellValue("SCHEDULE_STATUS")
                     buyerCo = GetExcelCellValue("BUYER_COMPANY")
                     If StrComp(scheduleStatus, "Updated", vbTextCompare) = 0 _
                        And StrComp(buyerCo, "BOEING COMMERCIAL AIRPLANES", vbTextCompare) = 0 Then
                        'And StrComp(buyerCo, "THE BOEING COMPANY", vbTextCompare) = 0 Then
                        strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
                        ImportExcelHeaderData Fields, "CHG"
                        GetExcelItemDetail strBuyerOrderNum, "CHG"
                     End If
                End If
            End If
        End If
         
        If (Fields(0) = "") Then
           bContinue = False
        End If
        
        'If iIndex > 200 Then bContinue = False ' use a subset to make debugging quicker
        
'         If iIndex Mod 500 = 0 Then
'            MsgBox CStr(iIndex)
'            Debug.Print CStr(iIndex)
'         End If
         
         iIndex = iIndex + 1
      Wend
      

      
'      While (bContinue)
'
'        ReadAllFields iIndex, ws
'
'        If (Fields(0) <> "") Then
'            strOrderType = GetExcelCellValue("ORDER_TYPE")
'
'            If (bNewImport = True) Then
'
'                If (strOrderType <> "Updated") Then
'
'                    strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
'
'                    ImportExcelHeaderData Fields, "NewOrder"
'
'                    GetExcelItemDetail strBuyerOrderNum, "NewOrder"
'                End If
'            Else
'
'                If (strOrderType = "Updated") Then
'                     ' 4/5/2020: Only schedule status = "Updated" also
'                     ' Buyer Company Region = "Boeing Commercial Airplanes"
'                     ' Buyer Company = "The Boeing Company"
'                     Dim scheduleStatus As String, buyerCo As String
'                     scheduleStatus = GetExcelCellValue("SCHEDULE_STATUS")
'                     'buyerRegion = GetExcelCellValue("BUYER_REGION")
'                     buyerCo = GetExcelCellValue("BUYER_COMPANY")
'                     If StrComp(scheduleStatus, "Updated", vbTextCompare) = 0 _
'                        And StrComp(buyerRegion, "BOEING COMMERCIAL AIRPLANES", vbTextCompare) = 0 Then
'                        'And StrComp(buyerCo, "THE BOEING COMPANY", vbTextCompare) = 0 Then
'                        strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
'                        ImportExcelHeaderData Fields, "CHG"
'                        GetExcelItemDetail strBuyerOrderNum, "CHG"
'                     End If
'                End If
'            End If
'        End If
'
'        If (Fields(0) = "") Then
'           bContinue = False
'        End If
'
'        If iIndex > 200 Then bContinue = False
'
''         If iIndex Mod 500 = 0 Then
''            MsgBox CStr(iIndex)
''            Debug.Print CStr(iIndex)
''         End If
'
'         iIndex = iIndex + 1
'      Wend
'
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
   End If
   'StopwatchStop "ParseExcelFile"
   Exit Function
   
DiaErr1:
   sProcName = "ParseExcelFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Function ReadAllFields(ByVal iIndex As Integer, ByRef ws As Worksheet)

    Dim iCols As Integer
        
    Fields(0) = ""
    While (iCols < 150)
        Fields(iCols) = ""
        iCols = iCols + 1
    Wend
    
    iCols = 0
    If (iIndex > 0 And Not ws Is Nothing) Then
        
        While (iCols < 150)
            Fields(iCols) = ws.Cells(iIndex, iCols + 1)
            iCols = iCols + 1
        Wend
    End If

End Function

'Function ParseCVSDFile(ByVal strFilePath As String) As Integer
'
'    MouseCursor ccHourglass
'    Grd.Rows = 1
'    On Error GoTo DiaErr1
'
'    ' Read the content if the text file.
'    Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
'    Dim lngPos As Integer
'    Dim bFound As Boolean
'    Dim strBuyerOrderNum As String
'    Dim strOrderType As String
'
'    ' Get a free file number
'    nFileNum = FreeFile
'
'    Open strFilePath For Input As nFileNum
'    ' Read the contents of the file
'    bFound = False
'    Dim iret As Integer
'    Line Input #nFileNum, sNextLine
'    Do While Not EOF(nFileNum)
'        Line Input #nFileNum, sNextLine
'
'        sNextLine = RemoveCommas(sNextLine)
'        Fields = Split(sNextLine, ",")
'
'        strOrderType = GetExcelCellValue("ORDER_TYPE")
'
'        If (bNewImport = True) Then
'
'            If (strOrderType <> "ChangeToPurchaseOrder") Then
'
'                strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
'
'                ImportExcelHeaderData Fields, "NewOrder"
'
'                GetExcelItemDetail strBuyerOrderNum, "NewOrder"
'            End If
'        Else
'
'            If (strOrderType = "ChangeToPurchaseOrder") Then
'                strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
'
'                ImportExcelHeaderData Fields, "ChngOrder"
'
'                GetExcelItemDetail strBuyerOrderNum, "ChngOrder"
'            End If
'        End If
'
'   Loop
'   Close nFileNum
'
'Exit Function
'
'
'   MouseCursor ccArrow
'
'   Exit Function
'
'DiaErr1:
'   sProcName = "fillgrid"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function
'
Function RemoveCommas(sNextLine As String) As String
    
    Dim length As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStrip As String
    lngStart = 1
    lngStart = InStr(lngStart, sNextLine, """")
    'lngEnd = InStr(lngStart + 1, sNextLine, """")
    
    length = Len(sNextLine)

    'strStrip = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
    
    'RemoveCommas = strStrip
    
    While (lngStart > 0)
        lngEnd = InStr(lngStart + 1, sNextLine, """")
        If (lngEnd > 0) Then
            'ReplaceComma sNextLine, lngStart, lngEnd
            sNextLine = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
        End If
        lngStart = InStr(1, sNextLine, """")
    Wend
    
    

End Function

Function ReplaceComma(sNextLine As String, lngStart As Long, lngEnd As Long)
    Dim i As Long
    i = lngStart
    While ((i <= lngEnd) And i > 0)
        i = InStr(i, sNextLine, ",")
        If (i > 0 And i <= lngEnd) Then
            sNextLine = Replace(sNextLine, ",", "-", i, 1)
            i = i + 1
        End If
    Wend

End Function

Function ImportExcelHeaderData(ByRef Fields() As String, strChgType As String)
   
   On Error GoTo DiaErr1
   
   Dim strContactName As String
   Dim strContactNum As String
   Dim strContactType As String
   
   Dim strShpToAddrs1 As String
   Dim strShpToPtyName1 As String
   Dim strShpToPtyName2 As String
   Dim strShpToPtyStreet As String
   Dim strShpToPtyStreetSup1 As String
   
   Dim strShpToPtyPostalCode As String
   Dim strShpToPtyCity As String
   Dim strShpToPtyRegionCode As String
   Dim strShpToPtyStreetSup2 As String
   Dim strShpToPtyNewAddress As String
   
   Dim strShpToName As String
   Dim strShpToLocName1 As String
   Dim strShpToLocName2 As String
   Dim strShpToLocStreet As String
   Dim strShpToLocStreetSup1 As String
   Dim strShpToLocPostalCode As String
   Dim strShpToLocCity As String
   Dim strShpToLocRegionCode As String
   Dim strShpToLocStreetSup2 As String
   Dim strShpToLocNewAddress As String
   
   Dim strNewAddress As String
   Dim strBuyerPartyAddrName As String
   Dim strBuyerOrderNum As String
   Dim iret As Integer
   
   
   ' Buyer Order Number
   strBuyerOrderNum = GetExcelCellValue("SOPO_BUYERORDNUM")
   
   ' Buyer Party address name
   strBuyerPartyAddrName = GetExcelCellValue("BUYER_PARTY_ADDR_NAME1")
   
   ' Buyer Party Ship name
   strShpToPtyName1 = GetExcelCellValue("SHPTO_PARTY_ADDRNAME1")
   
   strContactName = GetExcelCellValue("CONTACT_NAME")
   
   strContactNum = GetExcelCellValue("CONTACT_NUMBER")
   
   'strContactNum = Format(Trim(strContactNum), "###-###-####")
   '@@@ strContactType = GetExcelCellValue("CONTACT_TYPE")
   
   strShpToPtyName1 = GetExcelCellValue("SHPTO_PARTY_ADDRNAME1")
   
   strShpToPtyName2 = GetExcelCellValue("SHPTO_PARTY_ADDRNAME2")
   
   strShpToPtyStreet = GetExcelCellValue("SHPTO_PARTY_ADDRSTREET")
   
   
   strShpToPtyStreetSup1 = GetExcelCellValue("SHPTO_PARTY_ADDRSTRSUP1")
   strShpToPtyStreetSup2 = GetExcelCellValue("SHPTO_PARTY_ADDRSTRSUP2")
   
   strShpToPtyPostalCode = GetExcelCellValue("SHPTO_PARTY_POSTCODE")
   
   
   strShpToPtyCity = GetExcelCellValue("SHPTO_PARTY_CITY")
   
   strShpToPtyRegionCode = GetExcelCellValue("SHPTO_PARTY_REGCODE")
   
   MakeAddress strShpToPtyName2, strShpToPtyStreet, strShpToPtyStreetSup1, _
            strShpToPtyStreetSup2, strShpToPtyCity, strShpToPtyRegionCode, _
            strShpToPtyPostalCode, strShpToPtyNewAddress
    
    
   strShpToLocName1 = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRNAME1")
   strShpToLocName2 = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRNAME2")
   
   strShpToLocStreet = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRSTREET")
   strShpToPtyStreetSup1 = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRSTRSUP1")
   strShpToLocStreetSup2 = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRSTRSUP2")
   strShpToLocPostalCode = GetExcelCellValue("ORDDET_SHPTO_LOC_POSTCODE")
   strShpToLocCity = GetExcelCellValue("ORDDET_SHPTO_LOC_CITY")
   strShpToLocRegionCode = GetExcelCellValue("ORDDET_SHPTO_LOC_REGCODE")
    
    
   MakeAddress strShpToLocName2, strShpToLocStreet, strShpToPtyStreetSup1, _
            strShpToLocStreetSup2, strShpToLocCity, strShpToLocRegionCode, _
            strShpToLocPostalCode, strShpToLocNewAddress
   
   
   If (strShpToLocNewAddress <> "") Then
      strNewAddress = strShpToLocNewAddress
      strShpToName = strShpToLocName1
   Else
      strNewAddress = strShpToPtyNewAddress
      strShpToName = strShpToPtyName1
   End If
   
   Dim strSORemark As String
   Dim strContractNum As String
   Dim strRefTypeCode As String
   Dim strRefNum As String
   
   strSORemark = ""
   strContractNum = ""
   strRefNum = ""
   

   sSql = "INSERT INTO ExohdImport (SOPO_BUYERORDNUM, EXOSTART_IMPORT_TYPE, " _
            & "BUYER_PARTY_ADDR_NAME1, CONTACT_NAME," _
            & "CONTACT_NUMBER, CONTACT_TYPE, SHPTO_PARTY_ADDRNAME1, SHPTO_PARTY_ADDRNAME2, " _
            & "SHPTO_PARTY_ADDRSTREET, SHPTO_PARTY_ADDRSTRSUP1," _
            & "SHPTO_PARTY_ADDRSTRSUP2, SHPTO_PARTY_POSTCODE, " _
            & "SHPTO_PARTY_CITY, SHPTO_PARTY_REGCODE, " _
            & "ORDDET_SHPTO_LOC_ADDRNAME1,ORDDET_SHPTO_LOC_ADDRNAME2," _
            & "ORDDET_SHPTO_LOC_ADDRSTREET, ORDDET_SHPTO_LOC_ADDRSTRSUP1, " _
            & "ORDDET_SHPTO_LOC_ADDRSTRSUP2,ORDDET_SHPTO_LOC_POSTCODE, " _
            & "ORDDET_SHPTO_LOC_CITY, ORDDET_SHPTO_LOC_REGCODE, REFTYPE_CODE, " _
            & "REF_PRIMARY_REFNUM,REF_SUPPORT_REFNUM , REF_DESCRIPTION)" _
      & "VALUES('" & strBuyerOrderNum & "','" & strChgType & "','" _
            & strBuyerPartyAddrName & "','" & strContactName & "','" _
            & strContactNum & "','" & strContactType & "','" _
            & strShpToPtyName1 & "','" & strShpToPtyName2 & "','" _
            & strShpToPtyStreet & "','" & strShpToPtyStreetSup1 & "','" _
            & strShpToPtyStreetSup2 & "','" & strShpToPtyPostalCode & "','" _
            & strShpToPtyCity & "','" & strShpToPtyRegionCode & "','" _
            & strShpToLocName1 & "','" & strShpToLocName2 & "','" _
            & strShpToLocStreet & "','" & strShpToPtyStreetSup1 & "','" _
            & strShpToLocStreetSup2 & "','" & strShpToLocPostalCode & "','" _
            & strShpToLocCity & "','" & strShpToLocRegionCode & "','" _
            & strRefTypeCode & "','" & strContractNum & "','" _
            & strRefNum & "','" & strSORemark & "')"

   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   ' Insert to the database
   Exit Function
   
DiaErr1:
   sProcName = "ImportHeaderData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function



Function ImportHeaderData(ByRef ChgOrderNode As MSXML2.IXMLDOMNode, ByVal strChgType As String, ByVal strPrePath As String)
   
   On Error GoTo DiaErr1
   
   Dim strContactName As String
   Dim strContactNum As String
   Dim strContactType As String
   
   Dim strShpToAddrs1 As String
   Dim strShpToPtyName1 As String
   Dim strShpToPtyName2 As String
   Dim strShpToPtyStreet As String
   Dim strShpToPtyStreetSup1 As String
   
   Dim strShpToPtyPostalCode As String
   Dim strShpToPtyCity As String
   Dim strShpToPtyRegionCode As String
   Dim strShpToPtyStreetSup2 As String
   Dim strShpToPtyNewAddress As String
   
   Dim strShpToName As String
   Dim strShpToLocName1 As String
   Dim strShpToLocName2 As String
   Dim strShpToLocStreet As String
   Dim strShpToLocStreetSup1 As String
   Dim strShpToLocPostalCode As String
   Dim strShpToLocCity As String
   Dim strShpToLocRegionCode As String
   Dim strShpToLocStreetSup2 As String
   Dim strShpToLocNewAddress As String
   
   Dim strNewAddress As String
   Dim strBuyerPartyAddrName As String
   Dim strBuyerOrderNum As String
   
   ' Buyer Order Number
   GetNodeText ChgOrderNode, strBuyerOrderNum, GetNodeElemPath("SOPO_BUYERORDNUM", strPrePath)
   
   ' Buyer Party address name
   GetNodeText ChgOrderNode, strBuyerPartyAddrName, GetNodeElemPath("BUYER_PARTY_ADDR_NAME1", strPrePath)
   'strBuyerPartyAddrName = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/NameAddress/Name1").Text

   ' Buyer Party Ship name
   GetNodeText ChgOrderNode, strShpToPtyName1, GetNodeElemPath("SHPTO_PARTY_ADDRNAME1", strPrePath)
   'strShpToPtyName1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name1").Text
   
   GetNodeText ChgOrderNode, strContactName, GetNodeElemPath("CONTACT_NAME", strPrePath)
   GetNodeText ChgOrderNode, strContactNum, GetNodeElemPath("CONTACT_NUMBER", strPrePath)
   strContactNum = Format(Trim(strContactNum), "###-###-####")
   
   GetNodeText ChgOrderNode, strContactType, GetNodeElemPath("CONTACT_TYPE", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyName1, GetNodeElemPath("SHPTO_PARTY_ADDRNAME1", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyName2, GetNodeElemPath("SHPTO_PARTY_ADDRNAME2", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyStreet, GetNodeElemPath("SHPTO_PARTY_ADDRSTREET", strPrePath)
   
   'strStreetSup1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement2").Text
   GetNodeText ChgOrderNode, strShpToPtyStreetSup1, GetNodeElemPath("SHPTO_PARTY_ADDRSTRSUP1", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyStreetSup2, GetNodeElemPath("SHPTO_PARTY_ADDRSTRSUP2", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyPostalCode, GetNodeElemPath("SHPTO_PARTY_POSTCODE", strPrePath)
   
   GetNodeText ChgOrderNode, strShpToPtyCity, GetNodeElemPath("SHPTO_PARTY_CITY", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyRegionCode, GetNodeElemPath("SHPTO_PARTY_REGCODE", strPrePath)
   
   MakeAddress strShpToPtyName2, strShpToPtyStreet, strShpToPtyStreetSup1, _
            strShpToPtyStreetSup2, strShpToPtyCity, strShpToPtyRegionCode, _
            strShpToPtyPostalCode, strShpToPtyNewAddress
    
    
   GetNodeText ChgOrderNode, strShpToLocName1, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRNAME1", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocName2, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRNAME2", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocStreet, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRSTREET", strPrePath)
   GetNodeText ChgOrderNode, strShpToPtyStreetSup1, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRSTRSUP1", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocStreetSup2, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRSTRSUP2", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocPostalCode, GetNodeElemPath("ORDDET_SHPTO_LOC_POSTCODE", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocCity, GetNodeElemPath("ORDDET_SHPTO_LOC_CITY", strPrePath)
   GetNodeText ChgOrderNode, strShpToLocRegionCode, GetNodeElemPath("ORDDET_SHPTO_LOC_REGCODE", strPrePath)
    
   MakeAddress strShpToLocName2, strShpToLocStreet, strShpToPtyStreetSup1, _
            strShpToLocStreetSup2, strShpToLocCity, strShpToLocRegionCode, _
            strShpToLocPostalCode, strShpToLocNewAddress
   
   
   If (strShpToLocNewAddress <> "") Then
      strNewAddress = strShpToLocNewAddress
      strShpToName = strShpToLocName1
   Else
      strNewAddress = strShpToPtyNewAddress
      strShpToName = strShpToPtyName1
   End If
   
   Dim RefListNode As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
   Dim ReferenceCodeNode As MSXML2.IXMLDOMNode
   Dim strSORemark As String
   Dim strContractNum As String
   Dim strRefTypeCode As String
   Dim strRefNum As String
   
   
   If (strChgType = "ORG") Then
      Set RefListNode = ChgOrderNode.selectSingleNode(GetNodeElemPath("CHG_ORG_LISTOF_REF_CODE", "./"))
   Else
      Set RefListNode = ChgOrderNode.selectSingleNode(GetNodeElemPath("CHG_LISTOF_REF_CODE", "./"))
   End If
   
   strSORemark = ""
   If (Not RefListNode Is Nothing) Then
      For Each ReferenceCodeNode In RefListNode.childNodes
         
         GetNodeText ReferenceCodeNode, strRefTypeCode, GetNodeElemPath("REFTYPE_CODE", "./")
         
         If (strRefTypeCode = "ContractNumber") Then
            GetNodeText ReferenceCodeNode, strContractNum, GetNodeElemPath("REF_PRIMARY_REFNUM", "./")
            GetNodeText ReferenceCodeNode, strRefNum, GetNodeElemPath("REF_SUPPORT_REFNUM", "./")
            'strSORemark = strContractNum & " " & strRefNum
         End If
         
         
         If (strRefTypeCode = "LettersorNotes") Then
            GetNodeText ReferenceCodeNode, strRefNum, GetNodeElemPath("REF_PRIMARY_REFNUM", "./")
            
            If (strRefNum = "Line Text") Then
               Dim tmpSORemark As String
               GetNodeText ReferenceCodeNode, tmpSORemark, GetNodeElemPath("REF_PRIMARY_REFNUM", "./")
               strSORemark = strSORemark + tmpSORemark
            End If
            
         End If
      Next
   End If
'         AddSalesOrder strNewSO, strBuyerOrderNumber, strContactName, strContactNum, _
'                           strBuyerPartyAddressName, strCusName, strNewAddress, strSoType
   ' Remove the any single quote
   strSORemark = Replace(strSORemark, Chr$(39), " ")  'single quote

   sSql = "INSERT INTO ExohdImport (SOPO_BUYERORDNUM, EXOSTART_IMPORT_TYPE, " _
            & "BUYER_PARTY_ADDR_NAME1, CONTACT_NAME," _
            & "CONTACT_NUMBER, CONTACT_TYPE, SHPTO_PARTY_ADDRNAME1, SHPTO_PARTY_ADDRNAME2, " _
            & "SHPTO_PARTY_ADDRSTREET, SHPTO_PARTY_ADDRSTRSUP1," _
            & "SHPTO_PARTY_ADDRSTRSUP2, SHPTO_PARTY_POSTCODE, " _
            & "SHPTO_PARTY_CITY, SHPTO_PARTY_REGCODE, " _
            & "ORDDET_SHPTO_LOC_ADDRNAME1,ORDDET_SHPTO_LOC_ADDRNAME2," _
            & "ORDDET_SHPTO_LOC_ADDRSTREET, ORDDET_SHPTO_LOC_ADDRSTRSUP1, " _
            & "ORDDET_SHPTO_LOC_ADDRSTRSUP2,ORDDET_SHPTO_LOC_POSTCODE, " _
            & "ORDDET_SHPTO_LOC_CITY, ORDDET_SHPTO_LOC_REGCODE, REFTYPE_CODE, " _
            & "REF_PRIMARY_REFNUM,REF_SUPPORT_REFNUM , REF_DESCRIPTION)" _
      & "VALUES('" & strBuyerOrderNum & "','" & strChgType & "','" _
            & strBuyerPartyAddrName & "','" & strContactName & "','" _
            & strContactNum & "','" & strContactType & "','" _
            & strShpToPtyName1 & "','" & strShpToPtyName2 & "','" _
            & strShpToPtyStreet & "','" & strShpToPtyStreetSup1 & "','" _
            & strShpToPtyStreetSup2 & "','" & strShpToPtyPostalCode & "','" _
            & strShpToPtyCity & "','" & strShpToPtyRegionCode & "','" _
            & strShpToLocName1 & "','" & strShpToLocName2 & "','" _
            & strShpToLocStreet & "','" & strShpToPtyStreetSup1 & "','" _
            & strShpToLocStreetSup2 & "','" & strShpToLocPostalCode & "','" _
            & strShpToLocCity & "','" & strShpToLocRegionCode & "','" _
            & strRefTypeCode & "','" & strContractNum & "','" _
            & strRefNum & "','" & strSORemark & "')"

   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   ' Insert to the database
   Exit Function
   
DiaErr1:
   sProcName = "ImportHeaderData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Function CreateSOFromXMLDataEx()

   On Error GoTo DiaErr1
 
   Dim strShipAddr As String
   Dim strBuyItmNum As String
   
   Dim strPartID As String
   Dim strTotQty As String
   Dim strQty As String
   Dim strUnitPrice As String
   Dim strRequestDlvDate As String
   Dim strReqDt As String
   Dim strYear, strMonth, strDay As String
   Dim strNodePath As String
   Dim strNodeVal As String
   Dim strIdent As String
   Dim strBuyerPartyAddressName As String
   
   Dim iList As Integer
   Dim strBuyerOrderNumber As String
   
   Dim strPrevBuyOrdNum As String
   Dim strSONumber As String
   
   Dim bIncRow As Boolean
   Dim iItem  As Integer
   Dim bDeliveryDate As Boolean
   Dim strCusName As String
   Dim bByte As Byte
   Dim strImpType As String
   
   strPrevBuyOrdNum = ""
      
   For iList = 1 To GrdChg.rows - 1 Step 2
      GrdChg.Col = 0
      GrdChg.Row = iList
      
      ' Only if the part is checked
      If GrdChg.CellPicture = Chkyes.Picture Then
         
         GrdChg.Col = 1
         strSONumber = GrdChg.Text
         
         GrdChg.Col = 2
         strBuyerOrderNumber = GrdChg.Text
         
         GrdChg.Col = 3
         strBuyItmNum = GrdChg.Text
         
         GrdChg.Col = 6
         strQty = GrdChg.Text
         
         GrdChg.Col = 7
         strUnitPrice = GrdChg.Text
         
         GrdChg.Col = 8
         strReqDt = GrdChg.Text
         
         GrdChg.Col = 9
         strShipAddr = GrdChg.Text
   
         strCusName = cmbCst
         ' Get Customer PO
         bByte = CheckForCustomerPO(strCusName, strBuyerOrderNumber)

         ' if entereing a new item and po exists, inform the user
         If bByte = 1 And optExostar(0).Value = 1 Then
            bByte = MsgBox("The Customer PO Is In Use. Continue?", _
                 ES_YESQUESTION, Caption)
            If bByte = vbNo Then
               Exit Function
            End If
         End If
   
         If (strBuyerOrderNumber <> strPrevBuyOrdNum) Then
            ' Update the Address
            UpdateSOhdAddress strBuyerOrderNumber, strSONumber
            strPrevBuyOrdNum = strBuyerOrderNumber
         End If
         
         sSql = "SELECT a.SOPO_BUYERORDNUM, a.EXOSTART_IMPORT_TYPE, " _
                  & "ISNULL(INDEX_NUM, 0) as INDEX_NUM, BUYER_PARTY_ADDR_NAME1, " _
                  & "SHPTO_PARTY_ADDRNAME1, a.ORDDET_SHPTO_LOC_ADDRNAME1,BUYER_LINEITEM_NUM, PART_ID, TOT_QTY," _
                  & "UNIT_PRICE, REQ_DELDATE,SCHED_QTY_VALUE, SCHED_LINE_ID " _
         & " FROM ExohdImport a, ExoitImport b " _
                  & " WHERE a.SOPO_BUYERORDNUM = b.SOPO_BUYERORDNUM " _
                  & " AND a.EXOSTART_IMPORT_TYPE = b.EXOSTART_IMPORT_TYPE" _
                  & " AND a.SOPO_BUYERORDNUM ='" & strBuyerOrderNumber & "' " _
                  & " AND b.BUYER_LINEITEM_NUM = '" & strBuyItmNum & "'" _
                  & " AND a.EXOSTART_IMPORT_TYPE = 'CHG'"

         
'         sSql = "SELECT a.SOPO_BUYERORDNUM, a.EXOSTART_IMPORT_TYPE, " & vbCrLf _
'               & "ISNULL(INDEX_NUM, 0) as INDEX_NUM, BUYER_PARTY_ADDR_NAME1, " & vbCrLf _
'               & "SHPTO_PARTY_ADDRNAME1, a.ORDDET_SHPTO_LOC_ADDRNAME1,BUYER_LINEITEM_NUM, PART_ID, TOT_QTY," & vbCrLf _
'               & "UNIT_PRICE, REQ_DELDATE,SCHED_QTY_VALUE, SCHED_LINE_ID " _
'         & " FROM ExohdImport a" & vbCrLf _
'               & "join ExoitImport b on a.SOPO_BUYERORDNUM = b.SOPO_BUYERORDNUM " & vbCrLf _
'               & "AND a.EXOSTART_IMPORT_TYPE = b.EXOSTART_IMPORT_TYPE" & vbCrLf _
'               & "left join SohdTable so on so.SOPO = a.SOPO_BUYERORDNUM " & vbCrLf _
'               & "left join SoitTable si on si.ITSO = so.SONUMBER" & vbCrLf _
'               & "where a.SOPO_BUYERORDNUM ='" & strBuyerOrderNumber & "' " & vbCrLf _
'               & "AND b.BUYER_LINEITEM_NUM = '" & strBuyItmNum & "'" & vbCrLf _
'               & "AND a.EXOSTART_IMPORT_TYPE = 'CHG'" & vbCrLf _
'               & "AND BUYER_LINEITEM_NUM = it.ITNUMBER"
'
         
         'Debug.Print sSql
         
         Dim RdoExo As ADODB.Recordset
         clsADOCon.ADOErrNum = 0
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
         If bSqlRows Then
            With RdoExo
               strBuyerOrderNumber = !SOPO_BUYERORDNUM
               strImpType = !EXOSTART_IMPORT_TYPE
               strBuyerPartyAddressName = !BUYER_PARTY_ADDR_NAME1
               'strShpToPtyName1 = !SHPTO_PARTY_ADDRNAME1
               'strShpToLocName1 = !ORDDET_SHPTO_LOC_ADDRNAME1
               iItem = CInt(!INDEX_NUM)
               strBuyItmNum = !BUYER_LINEITEM_NUM
               strPartID = !PART_ID
               strTotQty = !TOT_QTY
               strQty = !SCHED_QTY_VALUE
               strUnitPrice = !UNIT_PRICE
               strRequestDlvDate = !REQ_DELDATE
               .Close
            End With
            Set RdoExo = Nothing
         End If
   
         Dim strSoNum As String
         Dim strNextItem As String
         Dim bRet As Boolean
         ' Get SoNumber and Next Item
         bRet = GetExistSO(strBuyerOrderNumber, strSoNum, strNextItem)
         
         If (bRet = True) Then
            AddSoItemEx strSoNum, strNextItem, strBuyItmNum, _
                        strPartID, strQty, strUnitPrice, strRequestDlvDate
                        
            ' update grey grid row to show updated values
            GrdChg.Row = iList + 1
            GrdChg.Col = 6
            GrdChg.Text = strQty
            GrdChg.Col = 7
            GrdChg.Text = strUnitPrice
            GrdChg.Col = 8
            GrdChg.Text = strReqDt
                                    
         End If
      End If
      
   Next
   MouseCursor ccArrow
   
   If optSORev.Value = vbChecked Then
      
      Dim strSoType As String
      GetSOType strSoNum, strSoType
      
      OptSoXml = vbChecked
      SaleSLe02a.Show
      SaleSLe02a.OptSoXml = vbChecked
      SaleSLe02a.SetFocus
      SaleSLe02a.cmbSon = strSoNum
      SaleSLe02a.cmbPre = strSoType
   End If
   
   Exit Function
 
 
DiaErr1:
   sProcName = "CreateSOFromXMLDataEx"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Function GetSOType(ByVal strSoNum As String, ByRef strSoType As String)
   On Error GoTo DiaErr1
   
   Dim RdoRpt As ADODB.Recordset
   
   sSql = "SELECT SOTYPE FROM sohdTable WHERE SONUMBER ='" & strSoNum & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      strSoType = Trim(RdoRpt!SOTYPE)
      ClearResultSet RdoRpt
   Else
      strSoType = ""
   End If
   Set RdoRpt = Nothing
   
   Exit Function

DiaErr1:
   sProcName = "CreateSOFromXMLDataEx"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Function CreateSOFromXMLData(ByVal strFilePath As String, ByVal strInputBuyerONum As String, _
                  ByVal strNewSO As String, ByVal strSoType As String, ByVal strCusName As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   
   Dim OrderDoc As MSXML2.DOMDocument40  'XML document object
       
   'Create a new document object
   Set OrderDoc = New MSXML2.DOMDocument40
   OrderDoc.preserveWhiteSpace = False
   OrderDoc.async = False
   OrderDoc.validateOnParse = False
   
   OrderDoc.resolveExternals = True
   'Remove the cached schema, we'll restore it later if needed
   Set OrderDoc.schemas = Nothing
   
       
   OrderDoc.loadXML (strXML)
   If OrderDoc.parseError.errorCode <> 0 Then
      MsgBox "Document could not be parsed:" & vbCrLf & OrderDoc.parseError.reason
      Exit Function
   End If
        
      'Update the book list shown in the Treeview to match any changes
      Dim OrderNode As MSXML2.IXMLDOMNode      'the "book" node
      
      Dim strContactName As String
      Dim strContactNum As String
      Dim strContactType As String
      
      Dim strShpToPtyName1 As String
      Dim strShpToPtyName2 As String
      Dim strShpToPtyStreet As String
      Dim strShpToPtyStreetSup1 As String
      
      Dim strShpToPtyPostalCode As String
      Dim strShpToPtyCity As String
      Dim strShpToPtyRegionCode As String
      Dim strShpToPtyStreetSup2 As String
      Dim strShpToPtyNewAddress As String
      
      Dim strShpToName As String
      Dim strShpToLocName1 As String
      Dim strShpToLocName2 As String
      Dim strShpToLocStreet As String
      Dim strShpToLocStreetSup1 As String
      Dim strShpToLocPostalCode As String
      Dim strShpToLocCity As String
      Dim strShpToLocRegionCode As String
      Dim strShpToLocStreetSup2 As String
      Dim strShpToLocNewAddress As String
      
      Dim strNewAddress As String
      
      Dim strPartID As String
      Dim strTotQty As String
      Dim strQty As String
      Dim strUOMQty As String
      Dim strRefTypeCode As String
      Dim strRefNum As String
      Dim strRefDesc As String
      Dim strUnitPrice As String
      Dim strRequestDlvDate As String
      Dim strReqDt As String
      Dim strYear, strMonth, strDay As String
      Dim strNodePath As String
      Dim strNodeVal As String
      Dim strIdent As String
      Dim strBuyerPartyAddressName As String
      
      Dim i As Integer
      Dim strBuyerOrderNumber As String
      Dim bIncRow As Boolean
      Dim iItem  As Integer
      Dim bDeliveryDate As Boolean

      
      On Error Resume Next
      For Each OrderNode In OrderDoc.documentElement.childNodes
         If OrderNode.baseName = "Order" Then
            strBuyerOrderNumber = OrderNode.selectSingleNode("./OrderHeader/OrderNumber/BuyerOrderNumber").Text
            If (strInputBuyerONum = strBuyerOrderNumber) Then
               Exit For
            End If
         End If
      Next
      
      If (Not OrderNode Is Nothing) Then
         
         'strContactName = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/OrderContact/Contact/ContactName").Text
         strNodePath = "./OrderHeader/OrderParty/BuyerParty/Party/OrderContact/Contact/ContactName"
         GetNodeText OrderNode, strContactName, strNodePath
         
         strContactNum = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/OrderContact/Contact/ListOfContactNumber/ContactNumber/ContactNumberValue").Text
         strContactNum = Format(Trim(strContactNum), "###-###-####")
         strContactType = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/OrderContact/Contact/ListOfContactNumber/ContactNumber/ContactNumberTypeCoded").Text
         
         strBuyerPartyAddressName = OrderNode.selectSingleNode("./OrderHeader/OrderParty/BuyerParty/Party/NameAddress/Name1").Text
         strShpToPtyName1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name1").Text
         strShpToPtyName2 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Name2").Text
         strShpToPtyStreet = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Street").Text
         'strStreetSup1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement2").Text
         strNodePath = "./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement1"
         GetNodeText OrderNode, strShpToPtyStreetSup1, strNodePath
         strShpToPtyStreetSup2 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement2").Text
         
         strShpToPtyPostalCode = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/PostalCode").Text
         strShpToPtyCity = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/City").Text
         strShpToPtyRegionCode = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/Region/RegionCoded").Text
         
         MakeAddress strShpToPtyName2, strShpToPtyStreet, strShpToPtyStreetSup1, _
                  strShpToPtyStreetSup2, strShpToPtyCity, strShpToPtyRegionCode, _
                  strShpToPtyPostalCode, strShpToPtyNewAddress
          
          
         strShpToLocName1 = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Name1").Text
         strShpToLocName2 = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Name2").Text
         strShpToLocStreet = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Street").Text
         'strStreetSup1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement2").Text
         strNodePath = "./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/StreetSupplement1"
         GetNodeText OrderNode, strShpToPtyStreetSup1, strNodePath
         strShpToLocStreetSup2 = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/StreetSupplement2").Text
         
         strShpToLocPostalCode = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/PostalCode").Text
         strShpToLocCity = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/City").Text
         strShpToLocRegionCode = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Region/RegionCoded").Text
          
         MakeAddress strShpToLocName2, strShpToLocStreet, strShpToPtyStreetSup1, _
                  strShpToLocStreetSup2, strShpToLocCity, strShpToLocRegionCode, _
                  strShpToLocPostalCode, strShpToLocNewAddress
         
         
         If (strShpToLocNewAddress <> "") Then
            strNewAddress = strShpToLocNewAddress
            strShpToName = strShpToLocName1
         Else
            strNewAddress = strShpToPtyNewAddress
            strShpToName = strShpToPtyName1
         End If
         
         Dim RefListNode As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
         Dim ReferenceCodeNode As MSXML2.IXMLDOMNode
         Dim strSORemark As String
         Dim strContractNum As String
         Dim strBuyerItemNum As String
         Set RefListNode = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/BaseItemDetail/ListOfItemReferences/ListOfReferenceCoded")
         
         strSORemark = ""
         For Each ReferenceCodeNode In RefListNode.childNodes
            strRefTypeCode = ReferenceCodeNode.selectSingleNode("./ReferenceTypeCoded").Text
            
            If (strRefTypeCode = "ContractNumber") Then
               strContractNum = ReferenceCodeNode.selectSingleNode("./PrimaryReference/Reference/RefNum").Text
               strRefNum = ReferenceCodeNode.selectSingleNode("./SupportingReference/Reference/RefNum").Text
               
               'strSORemark = strContractNum & " " & strRefNum
            End If
            
            
            If (strRefTypeCode = "LettersorNotes") Then
               strRefNum = ReferenceCodeNode.selectSingleNode("./PrimaryReference/Reference/RefNum").Text
               
               If (strRefNum = "Line Text") Then
                  strSORemark = strSORemark + ReferenceCodeNode.selectSingleNode("./ReferenceDescription").Text
               End If
               
            End If
         Next


'         AddSalesOrder strNewSO, strBuyerOrderNumber, strContactName, strContactNum, _
'                           strBuyerPartyAddressName, strCusName, strNewAddress, strSoType
         ' Remove the any single quote
         strSORemark = Replace(strSORemark, Chr$(39), " ")  'single quote
               
         AddSalesOrder strNewSO, strBuyerOrderNumber, strContactName, strContactNum, _
                           strShpToName, strCusName, strNewAddress, strSoType, strSORemark
         
         Dim ItemDetailsNode As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
         Dim ItemDetailNode As MSXML2.IXMLDOMNode
         
         iItem = 1
         Set ItemDetailsNode = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail")
         
         For Each ItemDetailNode In ItemDetailsNode.childNodes
         
            strPartID = ""
            strUnitPrice = ""
            strTotQty = ""
            strIdent = ""
            strBuyerItemNum = ""
                              
            strBuyerItemNum = ItemDetailNode.selectSingleNode("./BaseItemDetail/LineItemNum/BuyerLineItemNum").Text
            strPartID = ItemDetailNode.selectSingleNode("./BaseItemDetail/ItemIdentifiers/PartNumbers/BuyerPartNumber/PartNum/PartID").Text
            strTotQty = ItemDetailNode.selectSingleNode("./BaseItemDetail/TotalQuantity/Quantity/QuantityValue").Text
            strIdent = ItemDetailNode.selectSingleNode("./BaseItemDetail/ItemIdentifiers/CommodityCode/Identifier/Ident").Text
            strUnitPrice = ItemDetailNode.selectSingleNode("./PricingDetail/ListOfPrice/Price/UnitPrice/UnitPriceValue").Text
            
            
            Dim nodeListSchLine As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
            Dim nodeScheduleLine As MSXML2.IXMLDOMNode
            Set nodeListSchLine = ItemDetailNode.selectSingleNode("./DeliveryDetail/ListOfScheduleLine")
            
            For Each nodeListSchLine In nodeListSchLine.childNodes
               
               strRequestDlvDate = ""
               strQty = ""
               bDeliveryDate = False
               
               strQty = nodeListSchLine.selectSingleNode("./Quantity/QuantityValue").Text
               GetDeliveryDate nodeListSchLine, strRequestDlvDate, strReqDt, bDeliveryDate
               
               'Add sales Item
               AddSoItem strNewSO, CStr(iItem), strBuyerOrderNumber, strIdent, _
                           strPartID, strQty, strUnitPrice, strReqDt
               iItem = iItem + 1
            Next
            ' Not need
            'iItem = iItem + 1
            'strReqDt = Format(Mid$(strRequestDlvDate, 1, 8), "yy/mm/dd")
         Next ' Item Detail
      
      End If ' Order
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSOFromXMLData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function CreateSOFromExcelData(ByVal strFilePath As String, ByVal strInputBuyerONum As String, _
                  ByVal strNewSO As String, ByVal strSoType As String, ByVal strCusName As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
    Dim strContactName As String
    Dim strContactNum As String
    Dim strContactType As String
    
    Dim strShpToPtyName1 As String
    Dim strShpToPtyName2 As String
    Dim strShpToPtyStreet As String
    Dim strShpToPtyStreetSup1 As String
    
    Dim strShpToPtyPostalCode As String
    Dim strShpToPtyCity As String
    Dim strShpToPtyRegionCode As String
    Dim strShpToPtyStreetSup2 As String
    Dim strShpToPtyNewAddress As String
    
    Dim strShpToName As String
    Dim strShpToLocName1 As String
    Dim strShpToLocName2 As String
    Dim strShpToLocStreet As String
    Dim strShpToLocStreetSup1 As String
    Dim strShpToLocPostalCode As String
    Dim strShpToLocCity As String
    Dim strShpToLocRegionCode As String
    Dim strShpToLocStreetSup2 As String
    Dim strShpToLocNewAddress As String
    
    Dim strNewAddress As String
    
    Dim strPartID As String
    Dim strTotQty As String
    Dim strQty As String
    Dim strUOMQty As String
    Dim strRefTypeCode As String
    Dim strRefNum As String
    Dim strRefDesc As String
    Dim strUnitPrice As String
    Dim strRequestDlvDate As String
    Dim strReqDt As String
    Dim strYear, strMonth, strDay As String
    Dim strNodePath As String
    Dim strNodeVal As String
    Dim strIdent As String
    Dim strBuyerPartyAddressName As String
    Dim strNextItem As String
    
    Dim i As Integer
    Dim strBuyerOrderNumber As String
    Dim bIncRow As Boolean
    Dim iItem  As Integer
    Dim bDeliveryDate As Boolean

    iItem = 1
    sSql = "SELECT DISTINCT a.SOPO_BUYERORDNUM, a.EXOSTART_IMPORT_TYPE, " _
        & "BUYER_PARTY_ADDR_NAME1, SHPTO_PARTY_ADDRNAME1, a.CONTACT_NAME, a.CONTACT_NUMBER, a.CONTACT_TYPE," _
        & "a.SHPTO_PARTY_ADDRNAME1, a.SHPTO_PARTY_ADDRNAME2, a.SHPTO_PARTY_ADDRSTREET," _
        & "a.SHPTO_PARTY_ADDRSTRSUP1, a.SHPTO_PARTY_ADDRSTRSUP2, a.SHPTO_PARTY_POSTCODE," _
        & "a.SHPTO_PARTY_CITY, a.SHPTO_PARTY_REGCODE, a.ORDDET_SHPTO_LOC_ADDRNAME1," _
        & "a.ORDDET_SHPTO_LOC_ADDRNAME2, a.ORDDET_SHPTO_LOC_ADDRSTREET, " _
        & "a.ORDDET_SHPTO_LOC_ADDRSTRSUP1, a.ORDDET_SHPTO_LOC_ADDRSTRSUP2," _
        & "a.ORDDET_SHPTO_LOC_POSTCODE, a.ORDDET_SHPTO_LOC_CITY, a.ORDDET_SHPTO_LOC_REGCODE" _
    & " FROM ExohdImport a WHERE SOPO_BUYERORDNUM = '" & strInputBuyerONum & "' " _
    & " ORDER BY a.SOPO_BUYERORDNUM"
    
    'Debug.Print sSql
    
    Dim RdoExo As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
   
    If bSqlRows Then
       With RdoExo
        strBuyerOrderNumber = !SOPO_BUYERORDNUM
        strContactName = !CONTACT_NAME
        strContactNum = !CONTACT_NUMBER
        strContactType = !CONTACT_TYPE
        
        strBuyerPartyAddressName = !BUYER_PARTY_ADDR_NAME1
        strShpToPtyName1 = !SHPTO_PARTY_ADDRNAME1
        strShpToPtyName2 = !SHPTO_PARTY_ADDRNAME2
        strShpToPtyStreet = !SHPTO_PARTY_ADDRSTREET
        strShpToPtyStreetSup1 = !SHPTO_PARTY_ADDRSTRSUP1
        strShpToPtyStreetSup2 = !SHPTO_PARTY_ADDRSTRSUP2
        
        strShpToPtyPostalCode = !SHPTO_PARTY_POSTCODE
        strShpToPtyCity = !SHPTO_PARTY_CITY
        strShpToPtyRegionCode = !SHPTO_PARTY_REGCODE
        
        MakeAddress strShpToPtyName2, strShpToPtyStreet, strShpToPtyStreetSup1, _
                 strShpToPtyStreetSup2, strShpToPtyCity, strShpToPtyRegionCode, _
                 strShpToPtyPostalCode, strShpToPtyNewAddress
         
         
        strShpToLocName1 = !ORDDET_SHPTO_LOC_ADDRNAME1
        strShpToLocName2 = !ORDDET_SHPTO_LOC_ADDRNAME2
        strShpToLocStreet = !ORDDET_SHPTO_LOC_ADDRSTREET
        'strStreetSup1 = OrderNode.selectSingleNode("./OrderHeader/OrderParty/ShipToParty/Party/NameAddress/StreetSupplement2").Text
        strShpToPtyStreetSup1 = !ORDDET_SHPTO_LOC_ADDRSTRSUP1
        strShpToLocStreetSup2 = !ORDDET_SHPTO_LOC_ADDRSTRSUP2
        
        strShpToLocPostalCode = !ORDDET_SHPTO_LOC_POSTCODE
        strShpToLocCity = !ORDDET_SHPTO_LOC_CITY
        strShpToLocRegionCode = !ORDDET_SHPTO_LOC_REGCODE
         
        MakeAddress strShpToLocName2, strShpToLocStreet, strShpToPtyStreetSup1, _
                 strShpToLocStreetSup2, strShpToLocCity, strShpToLocRegionCode, _
                 strShpToLocPostalCode, strShpToLocNewAddress
        
        
        If (strShpToLocNewAddress <> "") Then
           strNewAddress = strShpToLocNewAddress
           strShpToName = strShpToLocName1
        Else
           strNewAddress = strShpToPtyNewAddress
           strShpToName = strShpToPtyName1
        End If
       
       
         Dim strSORemark As String
         Dim strContractNum As String
         Dim strBuyerItemNum As String
         
         strSORemark = ""
         strContractNum = ""
         strSORemark = Replace(strSORemark, Chr$(39), " ")  'single quote
               
         Dim strSoNum As String
         Dim bRet As Boolean
         
         ' Get SoNumber and Next Item
         bRet = GetExistSO(strBuyerOrderNumber, strNewSO, strNextItem)
         If (bRet = True) Then
            MsgBox "Has Existing SO Number : " & strNewSO, vbInformation, Caption
            If (strNextItem <> "") Then iItem = Val(strNextItem)
         Else
            AddSalesOrder strNewSO, strBuyerOrderNumber, strContactName, strContactNum, _
                              strShpToName, strCusName, strNewAddress, strSoType, strSORemark
         End If
        .Close
        End With
    End If

    sSql = "SELECT DISTINCT BUYER_LINEITEM_NUM, PART_ID, TOT_QTY,UNIT_PRICE, " _
        & " REQ_DELDATE,SCHED_QTY_VALUE, SCHED_LINE_ID , Convert(int, BUYER_LINEITEM_NUM) as buyernumber" _
        & " FROM ExoitImport WHERE SOPO_BUYERORDNUM  = '" & strBuyerOrderNumber & "'" _
        & " ORDER BY  buyernumber"
     
    'Debug.Print sSql
    
    Dim RdoDetail As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoDetail, ES_STATIC)
    
    If bSqlRows Then
        
      With RdoDetail
      While Not .EOF
     
        strPartID = ""
        strUnitPrice = ""
        strTotQty = ""
        strIdent = ""
        strBuyerItemNum = ""
                          
        strBuyerItemNum = !BUYER_LINEITEM_NUM
        strPartID = !PART_ID
        strTotQty = !TOT_QTY
        strIdent = !SCHED_LINE_ID
        strUnitPrice = !UNIT_PRICE
       
        strReqDt = !REQ_DELDATE
        strQty = !SCHED_QTY_VALUE
        
        'Add sales Item
        AddSoItem strNewSO, CStr(iItem), strBuyerOrderNumber, strIdent, _
                    strPartID, strQty, strUnitPrice, strReqDt
        .MoveNext
        iItem = iItem + 1
      Wend
      .Close
      End With
        
    End If

   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSOFromXMLData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetDeliveryDate(ByVal nodeItemDetail As MSXML2.IXMLDOMNode, _
         ByRef strRequestDlvDate As String, ByRef strReqDt As String, _
         ByRef bDeliveryDate As Boolean)
      
   Dim strXMLdate As String
   Dim strYear As String
   Dim strMonth As String
   Dim strDay As String
   
   Dim Node As MSXML2.IXMLDOMNode
   Dim nOtherDlvDate As MSXML2.IXMLDOMNode
   strXMLdate = ""
   
   Set Node = nodeItemDetail.selectSingleNode(GetNodeElemPath("REQ_DELDATE_NODE", "./"))
   'Set Node = nodeItemDetail.selectSingleNode("RequestedDeliveryDate")
   
   If Not Node Is Nothing Then
      strXMLdate = Node.Text
      bDeliveryDate = True
   Else
      ' get the next delivery date
      
      'Set nOtherDlvDate = nodeItemDetail.selectSingleNode("./ListOfOtherDeliveryDate")
      Set nOtherDlvDate = nodeItemDetail.selectSingleNode(GetNodeElemPath("LST_OTHER_DELDATE_NODE", "./"))
      
      If Not nOtherDlvDate Is Nothing Then
         'strXMLdate = nOtherDlvDate.selectSingleNode("./ListOfDateCoded/DateCoded/Date").Text
         GetNodeText nOtherDlvDate, strXMLdate, GetNodeElemPath("LST_OTHER_DATE", "./")
      End If
      
      bDeliveryDate = False
      Set nOtherDlvDate = Nothing
   End If
   Set Node = Nothing
   
   If (strXMLdate <> "") Then
      strRequestDlvDate = strXMLdate
      strYear = Mid$(strXMLdate, 1, 4)
      strMonth = Mid$(strXMLdate, 5, 2)
      strDay = Mid$(strXMLdate, 7, 2)
      strReqDt = strMonth & "/" & strDay & "/" & strYear
   Else
      strReqDt = ""
      strRequestDlvDate = ""
      bDeliveryDate = False
   End If
End Sub

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
   If OptSoXml.Value = vbUnchecked Then FormUnload
    'FormUnload
    Set SaleSLf12a = Nothing
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
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

Private Sub GrdChg_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      
      GrdChg.Col = 0
      
      If GrdChg.Row >= 1 Then
         GrdChg.Col = 4
         ' only if the data is change
         If (GrdChg.Text = "CHG") Then
            If GrdChg.CellPicture = Chkyes.Picture Then
               Set GrdChg.CellPicture = Chkno.Picture
            Else
               Set GrdChg.CellPicture = Chkyes.Picture
            End If
         End If
      End If
      
    End If
   

End Sub


Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.rows - 1
        Grd.Col = 0
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub


Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd.Col = 0
   If Grd.Row >= 1 Then
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
End Sub

Private Sub GrdChg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If GrdChg.Row >= 1 Then
      GrdChg.Col = 4
      ' only if the data is change
      If (GrdChg.Text = "CHG") Then
         GrdChg.Col = 0
         If GrdChg.CellPicture = Chkyes.Picture Then
            Set GrdChg.CellPicture = Chkno.Picture
         Else
            Set GrdChg.CellPicture = Chkyes.Picture
         End If
      End If
   End If
End Sub

Private Function CheckForCustomerPO(ByVal strCustomer As String, ByVal strPONum As String) As Byte
   On Error GoTo modErr1
   Dim RdoCpo As ADODB.Recordset
   If Trim(strPONum) = "" Then
      CheckForCustomerPO = 0
   Else
      sSql = "Qry_GetCustomerPo '" & Compress(strCustomer) _
             & "','" & Trim(strPONum) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpo, ES_FORWARD)
      If bSqlRows Then
         With RdoCpo
            CheckForCustomerPO = 1
            ClearResultSet RdoCpo
         End With
      End If
   End If
   Set RdoCpo = Nothing
   Exit Function
   
modErr1:
   sProcName = "CheckForCustomerPO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   CheckForCustomerPO = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Private Function GetExcelCellValue(strSOVar As String) As String
    On Error GoTo DiaErr1
    Dim strVal As String
    Dim iret As Integer
    
    iret = GetExcelCellNum(strSOVar)
    If (Val(iret) <> 0) Then
        strVal = Fields(iret - 1)
    Else
        strVal = ""
    End If
    
    GetExcelCellValue = strVal
    Exit Function
   
DiaErr1:
    sProcName = "GetExcelCellValue"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors MdiSect.ActiveForm

End Function

'Private Function GetExcelCellValue2(strSOVar As String, ws As Worksheet, iIndex As Integer) As String
'    On Error GoTo DiaErr1
'    Dim strVal As String
'    Dim iret As Integer
'
'    iret = GetExcelCellNum(strSOVar)
'    If (Val(iret) <> 0) Then
'        strVal = Fields(iret - 1)
'    Else
'        strVal = ""
'    End If
'
'    GetExcelCellValue = strVal
'    Exit Function
'
'DiaErr1:
'    sProcName = "GetExcelCellValue"
'    CurrError.Number = Err.Number
'    CurrError.Description = Err.Description
'    DoModuleErrors MdiSect.ActiveForm
'
'End Function

Private Sub GetExcelCellNumbers()
   If CellCount = 0 Then
      sSql = "select lower(isnull(SO_VARIABLE,'')) as CellName," & vbCrLf _
         & "isnull(EXCEL_CELLNO_NEW,0) as CellNoCreate," & vbCrLf _
         & "isnull(EXCEL_CELLNO_UPD,0) as CellNoUpdate" & vbCrLf _
         & "from ExoStarExcelMap" & vbCrLf _
         & "order by SO_VARIABLE "
      Dim rs As ADODB.Recordset, rows As Integer
      rows = 0
      bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_STATIC)  ' keyset returns record count at BOF
      If bSqlRows Then
         With rs
            CellCount = .RecordCount
            ReDim CellNames(CellCount)
            ReDim CellNoCreate(CellCount)
            ReDim cellNoUpdate(CellCount)
            While Not .EOF
               rows = rows + 1
               CellNames(rows) = !CellName
               CellNoCreate(rows) = !CellNoCreate
               cellNoUpdate(rows) = !cellNoUpdate
               .MoveNext
            Wend
         End With
      End If
      Set rs = Nothing
   End If
End Sub


Private Function GetExcelCellNum(strSOVar As String) As Integer
    Dim RdoNodePath As ADODB.Recordset
    On Error GoTo DiaErr1
    Dim strXMLType As String
    
    
    
    strXMLType = "SO_CREATE"
   
'    If (bNewImport) Then
'         sSql = "SELECT ISNULL(EXCEL_CELLNO_NEW, 0) CELLNO FROM ExoStarExcelMap WHERE SO_VARIABLE = '" & strSOVar _
'                    & "' AND XML_TYPE = '" & strXMLType & "'"
'    Else
'         sSql = "SELECT ISNULL(EXCEL_CELLNO_UPD,0) CELLNO FROM ExoStarExcelMap WHERE SO_VARIABLE = '" & strSOVar _
'                    & "' AND XML_TYPE = '" & strXMLType & "'"
'    End If
'
'    bSqlRows = clsADOCon.GetDataSet(sSql, RdoNodePath, ES_FORWARD)
'    If bSqlRows Then
'       With RdoNodePath
'          If Not IsNull(.Fields(0)) Then
'             GetExcelCellNum = Val(Trim(!CELLNO))
'          Else
'             GetExcelCellNum = 0
'          End If
'          ClearResultSet RdoNodePath
'       End With
'    Else
'        GetExcelCellNum = 0
'    End If
'    Set RdoNodePath = Nothing

   ' to speed up 4/6/2020 cell numbers are now stored in an array.  Not optimum but 1000x faster
   
   Dim i As Integer, limit As Integer, Key As String
   limit = UBound(CellNames)
   Key = LCase(strSOVar)
   GetExcelCellNum = 0
   
   For i = 1 To limit
'Debug.Print CStr(i) & ": " & Key & " : " & CellNames(i)
      If Key = CellNames(i) Then
         If (bNewImport) Then
            GetExcelCellNum = CellNoCreate(i)
         Else
            GetExcelCellNum = cellNoUpdate(i)
         End If
         Exit For
      End If
   Next i
   
    If GetExcelCellNum = 0 Then
      diagnoseMissingNumCount = diagnoseMissingNumCount + 1
      If diagnoseMissingNumCount <= 50 Then
         Debug.Print strSOVar & " has no cell number defined for " & IIf(bNewImport, "New", "Revised") & "SO"
      End If
   End If
    
    Exit Function
   
DiaErr1:
    sProcName = "GetExcelCellNum"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors MdiSect.ActiveForm

End Function


Private Function GetNodeElemPath(strSOVar As String, Optional strAppendPath As String = "") As String
   Dim RdoNodePath As ADODB.Recordset
   On Error GoTo DiaErr1
   Dim strXMLElePath As String
   Dim strXMLType As String
   
   strXMLElePath = ""
   strXMLType = "SO_CREATE"
   
   sSql = "SELECT XML_ELEMENT_PATH FROM ExoStarXMLMap WHERE SO_VARIABLE = '" & strSOVar _
                  & "' AND XML_TYPE = '" & strXMLType & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNodePath, ES_FORWARD)
   If bSqlRows Then
      With RdoNodePath
         If Not IsNull(.Fields(0)) Then
            strXMLElePath = strAppendPath & !XML_ELEMENT_PATH
         Else
            strXMLElePath = ""
         End If
         ClearResultSet RdoNodePath
      End With
   End If
   Set RdoNodePath = Nothing
   
   GetNodeElemPath = strXMLElePath
   Exit Function
   
DiaErr1:
   sProcName = "GetNodeElemPath"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function
Private Function GetNode(ByRef RootNode As MSXML2.IXMLDOMNode, _
            ByRef Node As MSXML2.IXMLDOMNode, ByVal strNodePath As String)

   On Error Resume Next
   Set Node = RootNode.selectSingleNode(strNodePath)
   
End Function
Private Function GetNodeText(ByRef OrderNode As MSXML2.IXMLDOMNode, _
                  ByRef strNodeVal As String, ByVal strNodePath As String)
         
   On Error Resume Next
   
   If (strNodePath <> "") Then
      
      Dim Node As MSXML2.IXMLDOMNode      'the "book" node
      strNodeVal = ""
      Set Node = OrderNode.selectSingleNode(strNodePath)
      
      If Not Node Is Nothing Then
         strNodeVal = Node.Text
      Else
         strNodeVal = ""
      End If
      
      Set Node = Nothing
   Else
         strNodeVal = ""
   End If
End Function


Private Function MakeAddress(strShipName2 As String, strStreet As String, strStreetSup1 As String, _
                  strStreetSup2 As String, strCity As String, strRegionCode As String, _
                  strPostalCode As String, ByRef strNewAddress As String)

   strNewAddress = ""
   
   ' MM not needed
   'If (strShipName2 <> "") Then strNewAddress = strNewAddress & strShipName2 & vbCrLf
   If (strStreet <> "") Then strNewAddress = strNewAddress & strStreet & vbCrLf
   If (strStreetSup1 <> "") Then strNewAddress = strNewAddress & strStreetSup1 & vbCrLf
   If (strStreetSup2 <> "") Then strNewAddress = strNewAddress & strStreetSup2 & vbCrLf
   
   ' moved Region ==> shiped
   'If (strRegionCode <> "") Then strNewAddress = strNewAddress & strRegionCode & vbCrLf
   If (strCity <> "") Then strNewAddress = strNewAddress & strCity
   
   If (strPostalCode <> "") Then
      If (strRegionCode <> "") Then
         strNewAddress = strNewAddress & ", " & IIf((strRegionCode <> ""), strRegionCode, "") & " - " & strPostalCode
      Else
         strNewAddress = strNewAddress & " - " & strPostalCode
      End If
   End If

End Function

Private Function ConvertXmlDateFormat(ByVal strInDate As String, ByRef strOutDate As String) As Boolean
   
   Dim strYear As String
   Dim strMonth As String
   Dim strDay As String
   
   If (strInDate <> "") Then
      strYear = Mid$(strInDate, 1, 4)
      strMonth = Mid$(strInDate, 5, 2)
      strDay = Mid$(strInDate, 7, 2)
      strOutDate = strMonth & "/" & strDay & "/" & strYear
      ConvertXmlDateFormat = True
   Else
      strOutDate = ""
      ConvertXmlDateFormat = False
   End If
End Function

Private Function CheckRecordExits(sSql As String)
    
   Dim RdoCon As ADODB.Recordset
   
   On Error GoTo ERR1
      
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
   If bSqlRows Then
       CheckRecordExits = True
   Else
       CheckRecordExits = False
   End If
   Set RdoCon = Nothing
   Exit Function
   
ERR1:
    CheckRecordExits = False

End Function

Private Function GetExcelItemDetail(ByVal strBuyerOrderNum As String, _
                        ByVal strChgType As String)

   On Error GoTo DiaErr1
   
   Dim strRequestDlvDate As String
   Dim strPartID As String
   Dim bDeliveryDate As Boolean
   Dim bPartFound As Boolean
   Dim strTotQty As String
   Dim strQty As String
   Dim strSchedID As String
   Dim strShpToLocName1 As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strShpToPtyName1 As String
   Dim strBuyLineItmNum As String
   Dim strItemIden As String
   Dim Index As Integer
   Index = 0
   strRequestDlvDate = ""
   strPartID = ""
   strUnitPrice = ""
   strTotQty = ""
   bDeliveryDate = False
   strReqDt = ""
   strShpToLocName1 = ""
   strShpToPtyName1 = ""
   
   strShpToLocName1 = GetExcelCellValue("ORDDET_SHPTO_LOC_ADDRNAME1")
   
   strPartID = GetExcelCellValue("PART_ID")
   
   strBuyLineItmNum = GetExcelCellValue("BUYER_LINEITEM_NUM")
   
   strTotQty = GetExcelCellValue("TOT_QTY")
   
   strUnitPrice = GetExcelCellValue("UNIT_PRICE")
   
   strReqDt = GetExcelCellValue("REQ_DELDATE")
   
   strQty = GetExcelCellValue("SCHED_QTY_VALUE")
    
   strSchedID = GetExcelCellValue("SCHED_LINE_ID")
   
   sSql = "INSERT INTO ExoitImport (SOPO_BUYERORDNUM, EXOSTART_IMPORT_TYPE, " _
         & "ORDDET_SHPTO_LOC_ADDRNAME1, INDEX_NUM, BUYER_LINEITEM_NUM, PART_ID, TOT_QTY," _
         & "UNIT_PRICE, REQ_DELDATE, " _
         & "SCHED_QTY_VALUE, SCHED_LINE_ID)" _
      & "VALUES('" & strBuyerOrderNum & "','" & strChgType & "','" _
         & strShpToLocName1 & "'," & CStr(Index) & ",'" & strBuyLineItmNum & "','" _
         & strPartID & "','" & strTotQty & "','" & strUnitPrice & "','" _
         & strReqDt & "','" & strQty & "','" _
         & strSchedID & "')"
   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect


   Exit Function
   
DiaErr1:
   sProcName = "GetItemDetailNode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function

Private Function GetItemDetailNode(ByVal RootOrderNode As MSXML2.IXMLDOMNode, _
                     ByVal ItemDetailNode As MSXML2.IXMLDOMNode, _
                     ByVal strBuyerOrderNum As String, ByVal strChgType As String, _
                     Index As Integer)

   On Error GoTo DiaErr1
   
   Dim strRequestDlvDate As String
   Dim strPartID As String
   Dim bDeliveryDate As Boolean
   Dim bPartFound As Boolean
   Dim strTotQty As String
   Dim strQty As String
   Dim strSchedID As String
   Dim strShpToLocName1 As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strShpToPtyName1 As String
   Dim strBuyLineItmNum As String
   Dim strItemIden As String
   
   strRequestDlvDate = ""
   strPartID = ""
   strUnitPrice = ""
   strTotQty = ""
   bDeliveryDate = False
   strReqDt = ""
   strShpToLocName1 = ""
   strShpToPtyName1 = ""
   
   
   ' Buyer Party Ship name
   GetNodeText ItemDetailNode, strShpToLocName1, GetNodeElemPath("ORDDET_SHPTO_LOC_ADDRNAME1", "./")
   'strShpToLocName1 = OrderNode.selectSingleNode("./OrderDetail/ListOfItemDetail/ItemDetail/DeliveryDetail/ShipToLocation/Location/NameAddress/Name1").Text
   
   
   GetNodeText ItemDetailNode, strPartID, GetNodeElemPath("PART_ID", "./")
   'strPartID = ItemDetailNode.selectSingleNode("./BaseItemDetail/ItemIdentifiers/PartNumbers/BuyerPartNumber/PartNum/PartID").Text
   
   GetNodeText ItemDetailNode, strBuyLineItmNum, GetNodeElemPath("BUYER_LINEITEM_NUM", "./")
   
   'GetNodeText ItemDetailNode, strItemIden, GetNodeElemPath("ITEM_IDENTIFIER", "./")
   
   GetNodeText ItemDetailNode, strTotQty, GetNodeElemPath("TOT_QTY", "./")
   'strTotQty = ItemDetailNode.selectSingleNode("./BaseItemDetail/TotalQuantity/Quantity/QuantityValue").Text
   
   GetNodeText ItemDetailNode, strUnitPrice, GetNodeElemPath("UNIT_PRICE", "./")
   'strUnitPrice = ItemDetailNode.selectSingleNode("./PricingDetail/ListOfPrice/Price/UnitPrice/UnitPriceValue").Text
   
   GetNodeText ItemDetailNode, strRequestDlvDate, GetNodeElemPath("REQ_DELDATE", "./")
   'strRequestDlvDate = ItemDetailNode.selectSingleNode("./DeliveryDetail/ListOfScheduleLine/ScheduleLine/RequestedDeliveryDate").Text
   ConvertXmlDateFormat strRequestDlvDate, strReqDt
   
   Dim nodeListSchLine As MSXML2.IXMLDOMNode     'reused node for author, title, etc. elements
   Dim nodeScheduleLine As MSXML2.IXMLDOMNode
   
   
   'Set nodeListSchLine = ItemDetailNode.selectSingleNode("./DeliveryDetail/ListOfScheduleLine")
   Set nodeListSchLine = ItemDetailNode.selectSingleNode(GetNodeElemPath("LST_SCHED_LINE", "./"))
   
   If (Not nodeListSchLine Is Nothing) Then
      For Each nodeListSchLine In nodeListSchLine.childNodes
      
         strQty = ""
         GetNodeText nodeListSchLine, strQty, GetNodeElemPath("SCHED_QTY_VALUE", "./")
         'strQty = nodeListSchLine.selectSingleNode("./Quantity/QuantityValue").Text
         
         GetNodeText nodeListSchLine, strSchedID, GetNodeElemPath("SCHED_LINE_ID", "./")
         'strSchedID = nodeListSchLine.selectSingleNode("./ScheduleLineID").Text
         
         GetDeliveryDate nodeListSchLine, strRequestDlvDate, strReqDt, bDeliveryDate
      Next
   End If
   
   sSql = "INSERT INTO ExoitImport (SOPO_BUYERORDNUM, EXOSTART_IMPORT_TYPE, " _
         & "ORDDET_SHPTO_LOC_ADDRNAME1, INDEX_NUM, BUYER_LINEITEM_NUM, PART_ID, TOT_QTY," _
         & "UNIT_PRICE, REQ_DELDATE, " _
         & "SCHED_QTY_VALUE, SCHED_LINE_ID)" _
      & "VALUES('" & strBuyerOrderNum & "','" & strChgType & "','" _
         & strShpToLocName1 & "'," & CStr(Index) & ",'" & strBuyLineItmNum & "','" _
         & strPartID & "','" & strTotQty & "','" & strUnitPrice & "','" _
         & strReqDt & "','" & strQty & "','" _
         & strSchedID & "')"
   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect

   Exit Function
   
DiaErr1:
   sProcName = "GetItemDetailNode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function

Private Function DeleteOldData(strTableName As String)

   If (strTableName <> "") Then
      sSql = "DELETE FROM " & strTableName
      clsADOCon.ExecuteSql sSql
   End If

End Function

Private Function UpdateSOhdAddress(strBuyerOrderNumber As String, strSoNum As String)

   On Error GoTo DiaErr1
   
   sSql = "SELECT SOPO_BUYERORDNUM, EXOSTART_IMPORT_TYPE, " _
            & "BUYER_PARTY_ADDR_NAME1, SHPTO_PARTY_ADDRNAME1, SHPTO_PARTY_ADDRNAME2, " _
            & "SHPTO_PARTY_ADDRSTRSUP1, SHPTO_PARTY_ADDRSTRSUP2,SHPTO_PARTY_POSTCODE, " _
            & "SHPTO_PARTY_CITY, SHPTO_PARTY_REGCODE, SHPTO_PARTY_ADDRSTREET," _
            & "ORDDET_SHPTO_LOC_ADDRNAME1, ORDDET_SHPTO_LOC_ADDRNAME2,ORDDET_SHPTO_LOC_ADDRSTREET, " _
            & "ORDDET_SHPTO_LOC_ADDRSTRSUP1, ORDDET_SHPTO_LOC_ADDRSTRSUP2,ORDDET_SHPTO_LOC_POSTCODE, " _
            & "ORDDET_SHPTO_LOC_CITY, ORDDET_SHPTO_LOC_REGCODE " _
   & " FROM ExohdImport WHERE SOPO_BUYERORDNUM = '" & strBuyerOrderNumber & "'" _
            & " AND EXOSTART_IMPORT_TYPE = 'CHG'"
   
   'Debug.Print sSql
   
   Dim strShpToPtyNewAddress As String
   Dim strShpToLocNewAddress As String
   Dim strNewAddress As String
   Dim strShpToName As String
   Dim strImpType As String
   
   
   Dim strShpToPtyName1 As String
   Dim strShpToPtyName2 As String
   Dim strShpToPtyStreet As String
   Dim strShpToPtyStreetSup1 As String
   Dim strShpToPtyPostalCode As String
   Dim strShpToPtyCity As String
   Dim strShpToPtyRegionCode As String
   Dim strShpToPtyStreetSup2 As String
   
   Dim strShpToLocName1 As String
   Dim strShpToLocName2 As String
   Dim strShpToLocStreet As String
   Dim strShpToLocStreetSup1 As String
   Dim strShpToLocPostalCode As String
   Dim strShpToLocCity As String
   Dim strShpToLocRegionCode As String
   Dim strShpToLocStreetSup2 As String
   
   Dim RdoExo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
   If bSqlRows Then
      With RdoExo
         strBuyerOrderNumber = !SOPO_BUYERORDNUM
         strImpType = !EXOSTART_IMPORT_TYPE
         strShpToPtyName1 = !SHPTO_PARTY_ADDRNAME1
         strShpToPtyName2 = !SHPTO_PARTY_ADDRNAME2
         strShpToPtyStreet = !SHPTO_PARTY_ADDRSTREET
         strShpToPtyStreetSup1 = !SHPTO_PARTY_ADDRSTRSUP1
         strShpToPtyStreetSup2 = !SHPTO_PARTY_ADDRSTRSUP2
         strShpToPtyPostalCode = !SHPTO_PARTY_POSTCODE
         strShpToPtyCity = !SHPTO_PARTY_CITY
         strShpToPtyRegionCode = !SHPTO_PARTY_REGCODE
   
         MakeAddress strShpToPtyName2, strShpToPtyStreet, strShpToPtyStreetSup1, _
                  strShpToPtyStreetSup2, strShpToPtyCity, strShpToPtyRegionCode, _
                  strShpToPtyPostalCode, strShpToPtyNewAddress
         
         
         strShpToLocName1 = !ORDDET_SHPTO_LOC_ADDRNAME1
         strShpToLocName2 = !ORDDET_SHPTO_LOC_ADDRNAME2
         strShpToLocStreet = !ORDDET_SHPTO_LOC_ADDRSTREET
         strShpToPtyStreetSup1 = !ORDDET_SHPTO_LOC_ADDRSTRSUP1
         strShpToLocStreetSup2 = !ORDDET_SHPTO_LOC_ADDRSTRSUP2
         strShpToLocPostalCode = !ORDDET_SHPTO_LOC_POSTCODE
         strShpToLocCity = !ORDDET_SHPTO_LOC_CITY
         strShpToLocRegionCode = !ORDDET_SHPTO_LOC_REGCODE
         
         MakeAddress strShpToLocName2, strShpToLocStreet, strShpToPtyStreetSup1, _
                  strShpToLocStreetSup2, strShpToLocCity, strShpToLocRegionCode, _
                  strShpToLocPostalCode, strShpToLocNewAddress
   
         .Close
      End With
      Set RdoExo = Nothing
   End If
   
   If (strShpToLocNewAddress <> "") Then
      strNewAddress = strShpToLocNewAddress
      strShpToName = strShpToLocName1
   Else
      strNewAddress = strShpToPtyNewAddress
      strShpToName = strShpToPtyName1
   End If


   sSql = "UPDATE SohdTable SET SOSTNAME='" & strShpToName & "'," _
          & "SOSTADR='" & strNewAddress & "'" _
          & " WHERE SONUMBER=" & Val(strSoNum) & ""
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql 'rdExecDirect

   Exit Function
   
DiaErr1:
   sProcName = "GetItemDetailNode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Private Sub AddSoItemEx(strSoNum As String, strItem As String, strCusItem As String, _
                     strPartID As String, strQty As String, _
                     strUnitPrice As String, strReqDt As String)
   
   
   Dim RdoSoit As ADODB.Recordset
   Dim strShiped As String
   Dim NewReqDt As Date
   Dim strNewReqDt As String
   

   If (Trim(strCusItem) = "") Then
      MsgBox "Cusomer Item number is empty.", _
            vbInformation, Caption
      Exit Sub
   End If
   
   strShiped = ""
   sSql = "SELECT DISTINCT ITSO, ITNUMBER, ISNULL(ITPSSHIPPED, 0)  ITPSSHIPPED FROM SoitTable WHERE " _
             & " ITSO = '" & strSoNum & "'" _
             & "  AND CONVERT(int,ITCUSTITEMNO) = '" & Val(strCusItem) & "'"
             
                         
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_FORWARD)
   If bSqlRows Then
   
      strShiped = RdoSoit!ITPSSHIPPED
      strItem = RdoSoit!ITNUMBER
      
      ClearResultSet RdoSoit
      Set RdoSoit = Nothing
      
      If (CInt(strShiped) = 1) Then
         Exit Sub
      Else
          
         ' get the latest revision of Item and them updated the part
         Dim strRev As String
         Dim lRemQty As Long
         
         GetLatesSoitRev strSoNum, strItem, Val(strQty), strRev, lRemQty
         
         
         NewReqDt = DateAdd("d", -iFrtDays, CDate(strReqDt))
         
         strNewReqDt = Format(NewReqDt, "mm/dd/yyyy")
         
         sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strNewReqDt & "'," _
                  & " ITCUSTREQ = '" & strReqDt & "',ITSCHEDDEL = '" & strReqDt & "'," _
                  & " ITDOLLARS = " & strUnitPrice & vbCrLf _
            & " WHERE ITSO = '" & strSoNum & "' AND ITNUMBER = '" & Val(strItem) & "'" _
                  & " AND ITREV = '" & strRev & "'"
                           
'         sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strReqDt & "'," _
'                  & " ITCUSTREQ = '" & strReqDt & "',ITSCHEDDEL = '" & strReqDt & "' " _
'                  & " WHERE ITSO = '" & strSoNum & "' AND CONVERT(int,ITCUSTITEMNO) = '" & Val(strCusItem) & "'"

      End If
             
      'Debug.Print sSql
      
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      
      'Not needed
      MsgBox "Updated Sales Order '" & strSoNum & "' and Item '" & strItem & "'.", vbExclamation, Caption
      
      Exit Sub
   End If
   
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   NewReqDt = DateAdd("d", -iFrtDays, CDate(strReqDt))
         
   strNewReqDt = Format(NewReqDt, "mm/dd/yyyy")
         
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITCUSTITEMNO, ITPART,ITQTY,ITSCHED,ITSCHEDDEL,ITBOOKDATE," _
          & "ITDOLLORIG, ITDOLLARS, ITUSER) " _
          & "VALUES(" & strSoNum & "," & strItem & ",'" & Val(strCusItem) & "','" _
          & Compress(strPartID) & "'," & Val(strQty) & ",'" & strNewReqDt & "','" & strReqDt & "','" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & CCur(strUnitPrice) & "','" _
          & CCur(strUnitPrice) & "','" & sInitials & "')"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   'Add commission if applicable.
'   If cmdCom.Enabled Then
     Dim Item As New ClassSoItem
     Dim bUserMsg As Boolean
     bUserMsg = False
     Item.InsertCommission CLng(strSoNum), CLng(strItem), "", ""
     Item.UpdateCommissions CLng(strSoNum), CLng(strItem), "", bUserMsg
 '  End If==
   
   clsADOCon.CommitTrans
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetLatesSoitRev(ByVal strNewSO As String, ByVal strItem As String, _
            ByVal lEdiQty As Long, ByRef strRev As String, ByRef lRemQty As Long)
            
   Dim RdoSoit As ADODB.Recordset
   Dim strShiped As String
   Dim itemQty As Long
   
   lRemQty = 0
   strShiped = ""
   sSql = "SELECT DISTINCT ITSO, ITNUMBER, ITREV, ITQTY, ISNULL(ITPSSHIPPED, 0)  ITPSSHIPPED FROM SoitTable WHERE " _
             & " ITSO = '" & strNewSO & "'" _
             & "  AND ITNUMBER = '" & Val(strItem) & "'" & vbCrLf & "order by ITREV"

   'Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_FORWARD)
   If bSqlRows Then
   
      With RdoSoit
      While Not .EOF
            strRev = !itrev
            itemQty = !ITQty
            strShiped = !ITPSSHIPPED
            lEdiQty = lEdiQty - itemQty
            ' remainder
            lRemQty = lEdiQty
         .MoveNext
      Wend
      .Close
      End With
      Set RdoSoit = Nothing
   End If
            
End Function


