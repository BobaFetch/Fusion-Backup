VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PackPSf09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Invoices EDI file"
   ClientHeight    =   9810
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   15045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Selection All"
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
      Left            =   13080
      TabIndex        =   18
      ToolTipText     =   " Select All"
      Top             =   3600
      Width           =   1920
   End
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select import"
      Top             =   10080
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   10080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGetInv 
      Caption         =   "Get Invoices"
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2145
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1800
      Width           =   1555
   End
   Begin VB.CommandButton cmdInv 
      Caption         =   "Create Invoice  file"
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
      Left            =   7200
      TabIndex        =   6
      ToolTipText     =   " Create Invoice file"
      Top             =   9840
      Visible         =   0   'False
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
      Left            =   13080
      TabIndex        =   7
      ToolTipText     =   " Clear the selection"
      Top             =   4320
      Width           =   1920
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
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
      FormDesignHeight=   9810
      FormDesignWidth =   15045
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   13920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   6015
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   3480
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   3
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice File Name"
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   17
      Top             =   10080
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Inv Date"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "** Part Not found in Fusion"
      Height          =   255
      Left            =   13080
      TabIndex        =   13
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "* Not Delivery Dates"
      Height          =   255
      Index           =   0
      Left            =   13080
      TabIndex        =   12
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7680
      Picture         =   "PackPSf09a.frx":07AE
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7680
      Picture         =   "PackPSf09a.frx":0B38
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "PackPSf09a"
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

Dim strSoNum As String
Dim strSOStnme As String
Dim strSOStadr As String
Dim strSOVia As String
Dim strSOTerms As String

Dim iDays As Integer
Dim iFrtDays As Integer
Dim iNetDays As Integer

Dim cDiscount As Currency

Dim arrValue() As Variant
Dim arrFieldName() As Variant

Dim strPartNum As String
Dim strPartCnt As String
Dim strPAUnit As String
Dim strPartInfo As String

Dim strPartNumFld As String
Dim strPartCntFld As String
Dim strPAUnitFld As String
Dim strPartInfoFld As String

Private txtKeyPress As New EsiKeyBd


Private Sub ClearArrays(ByVal iSize As Integer)
    Erase arrValue
    ReDim arrValue(0 To 96, iSize)
End Sub



Private Sub cmdCan_Click()
    'sLastPrefix = cmbPre
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

Private Sub cmdGetInv_Click()
    Dim strWindows As String
    Dim strAccFileName As String
    Dim strpathFilename As String

    On Error GoTo DiaErr1
    FillGrid

    Exit Sub
DiaErr1:
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors Me

End Sub

Private Sub cmdOpenDia_Click()
    fileDlg.Filter = "EDI File (*.edi) | *.edi"
    fileDlg.ShowOpen
    If fileDlg.filename = "" Then
        txtEdiFilePath.Text = ""
    Else
        txtEdiFilePath.Text = fileDlg.filename
    End If
End Sub

Private Sub cmdInv_Click()

    Dim strFileName As String
    Dim strDate As String
    Dim strCust As String

    strFileName = Trim(txtEdiFilePath.Text)

    'strFilePath = "C:\Development\FusionCode\EDI\USC EDI Files\USC EDI Files\ASNOUT1.EDI"

    Dim nFileNum As Integer, lLineCount As Long
    Dim strBlank As String

    strDate = txtDte.Text
    strCust = cmbCst.Text

    strFileName = txtEdiFilePath.Text

    If (Trim(strFileName) <> "") Then
        ' Open the file
        nFileNum = FreeFile
      Open strFileName For Output As nFileNum

        If EOF(nFileNum) Then
            GenerateINVFile nFileNum, strDate, strCust
        End If
        ' Close the file
        Close (nFileNum)
    Else
        MsgBox "Select filename to create ASN export file.", _
           vbInformation, Caption
    End If

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
'    cmbCst = CheckLen(cmbCst, 10)
'    FindCustomer Me, cmbCst, False
'    lblNotice.Visible = False

End Sub


Private Sub CmdSelAll_Click()
   
   Dim iList As Integer
   For iList = 1 To Grd.Rows - 1
       Grd.Col = 0
       Grd.Row = iList
       ' Only if the part is checked
       If Grd.CellPicture = Chkno.Picture Then
           Set Grd.CellPicture = Chkyes.Picture
       End If
   Next
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then

      'FillCustomers
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
              & " A.CUREF = b.CUREF AND PACCARDPART = 1"
      
      LoadComboBox cmbCst, -1
      AddComboStr cmbCst.hwnd, "" & Trim("ALL")
      cmbCst = "ALL"
      txtNme = "*** All Customer selected ***"
      
      'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
      
      txtDte = Format(ES_SYSDATE, "mm/dd/yy")
      
      'GetPackslip True
      bOnLoad = 0
   End If
   MouseCursor (0)

End Sub


Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hwnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub


Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST

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

        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Apply"
        .Col = 1
        .Text = "InvNo"
        .Col = 2
        .Text = "PO Number"
        .Col = 3
        .Text = "PartNumber"
        .Col = 4
        .Text = "Qty"
        .Col = 5
        .Text = "Unit Price"
        .Col = 6
        .Text = "Inv Total"
        .Col = 7
        .Text = "SO Date"
        .Col = 8
        .Text = "Inv Date"
        .Col = 9
        .Text = "ASN Num"
        .Col = 10
        .Text = "Customer"
        .Col = 11
        .Text = "Units"

        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 2300
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1100
        .ColWidth(10) = 1000
        .ColWidth(11) = 800
        
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns

    End With

   Call WheelHook(Me.hwnd)
   bOnLoad = 1

End Sub

Function GenerateINVFile(ByVal nFileNum As Integer, ByVal strDate As String, ByVal strCust As String) As Integer

    Dim RdoPs As ADODB.Recordset

    MouseCursor (ccHourglass)
    On Error GoTo DiaErr1

    Dim strPONumber As String
    Dim strPartNum As String
    Dim strPSNum As String
    Dim strInvNum  As String
    Dim strQty As String
    Dim strInvDate As String
    Dim strSODate As String
    Dim strContainer As String
    Dim strUnitPrice As String
    Dim strPAUnit As String
    Dim bPartFound, bIncRow As Boolean
    Dim strInvTot As String
    Dim iList As Integer


   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
      
         Grd.Col = 1
         strInvNum = Trim(Grd.Text)
         Grd.Col = 2
         strPONumber = Trim(Grd.Text)
         Grd.Col = 3
         strPartNum = Trim(Grd.Text)
         Grd.Col = 4
         strQty = Trim(Grd.Text)
         Grd.Col = 5
         strUnitPrice = Trim(Grd.Text)
         Grd.Col = 6
         strInvTot = Trim(Grd.Text)
         Grd.Col = 7
         strSODate = Trim(Grd.Text)
         Grd.Col = 8
         strInvDate = Trim(Grd.Text)
         Grd.Col = 9
         strContainer = Trim(Grd.Text)
         Grd.Col = 10
         strCust = Trim(Grd.Text)
         Grd.Col = 11
         strPAUnit = Trim(Grd.Text)
         
         ' Add Header detail
         Dim strHeader As String
         CreateHeader strCust, strInvNum, strPONumber, strSODate, _
               strInvDate, strContainer, strHeader
         
         ' Read the contents of the file
         If EOF(nFileNum) Then
            Print #nFileNum, strHeader
            Debug.Print (strHeader)
         End If
         
         ' Add Detail
         Dim strInvDetail As String
         CreateInvDetail strInvNum, strQty, strPAUnit, strUnitPrice, strPartNum, strInvDetail
         ' Read the contents of the file
         If EOF(nFileNum) Then
            Print #nFileNum, strInvDetail
            Debug.Print (strInvDetail)
         End If
         
         ' Add So
         Dim strSODetail As String
         CreateSODetail strInvNum, strQty, strInvTot, strSODetail
         ' Read the contents of the file
         If EOF(nFileNum) Then
            Print #nFileNum, strSODetail
            Debug.Print (strSODetail)
         End If
      End If
   
   Next
   Set RdoPs = Nothing
   MouseCursor (ccArrow)
   MsgBox "Invoices selected for date '" & strDate & "' is exported.", _
      vbInformation, Caption
   
   Exit Function

DiaErr1:
    sProcName = "GenerateINVFile"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me

End Function



Function FillGrid() As Integer

    Dim strSenderCode As String
    Dim RdoInv As ADODB.Recordset

    MouseCursor (ccHourglass)
    Grd.Rows = 1
    On Error GoTo DiaErr1


    Dim strCust, strPONumber, strPartNum As String
    Dim strPSNum As String
    Dim strInvNum  As String
    Dim strDate, strQty As String
    Dim strInvDate, strSODate As String
    Dim strContainer As String
    Dim strUnitPrice, strPAUnit As String
    Dim bPartFound, bIncRow As Boolean
    Dim strInvTot As String
    Dim iItem As Integer

    strDate = txtDte.Text
    strCust = cmbCst.Text

   If (strCust = "ALL") Then
      strCust = ""
   End If
   
   
   
'   sSql = "SELECT PIPART, INVNO, PSNUMBER, INVTOTAL, INVDATE, INVCUST, SOPO, " _
'             & "SODATE, PISELLPRICE,PIQTY , PSCONTAINER, PSSHIPNO, PAPRICE, PAUNITS" _
'             & "  FROM cihdTable, sohdtable, soitTable,pshdTable, psittable, Parttable" _
'             & "  WHERE INVDATE = '" & strDate & "'" _
'             & "    AND INVCUST LIKE '" & strCust & "%'" _
'             & "    AND INVSO = SONUMBER" _
'             & "    AND PSINVOICE = INVNO" _
'             & "    AND PSNUMBER = PIPACKSLIP" _
'             & "    AND PARTREF = PIPART"

   sSql = "SELECT PIPART, PARTNUM, INVNO, PSNUMBER, INVTOTAL, INVDATE, INVCUST, SOPO, " _
             & "SODATE, PISELLPRICE,PIQTY , PSCONTAINER, PSSHIPNO, PAPRICE, PAUNITS" _
             & "  FROM cihdTable, sohdtable, soitTable, pshdTable, psittable, Parttable" _
             & " WHERE INVDATE = '" & strDate & "'" _
             & " AND INVCUST LIKE '" & strCust & "%'" _
             & " AND ITPSNUMBER = PSNUMBER" _
             & " AND ITSO = SONUMBER" _
             & " AND PSINVOICE = INVNO" _
             & " AND PSNUMBER = PIPACKSLIP" _
             & " AND PARTREF = PIPART" _
             & " AND INVCUST IN " _
             & "   (SELECT DISTINCT a.CUREF FROM " _
             & "     ASNInfoTable a, custtable b WHERE " _
             & "     A.CUREF = b.CUREF AND PACCARDPART = 1)"

             '& "    AND INVCUST LIKE '" & strCust & "%'" _

    Debug.Print (sSql)

    bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, adOpenStatic)

    If bSqlRows Then
        With RdoInv
            While Not .EOF

                strInvNum = Trim(!invno)
                strPSNum = Trim(!PsNumber)
                strContainer = Trim(!PSCONTAINER)
                strPONumber = Trim(!SOPO)
                strPartNum = Trim(!PartNum)
                strQty = Trim(!PIQTY)
                strInvTot = Trim(!INVTOTAL)
                strInvDate = Trim(!INVDATE)
                strSODate = Trim(!SODATE)
                strCust = Trim(!INVCUST)
                strUnitPrice = Trim(!PAPRICE)
                strPAUnit = Trim(!PAUNITS)

                Grd.Rows = Grd.Rows + 1
                Grd.Row = Grd.Rows - 1
                bIncRow = False
                iItem = 1

                Grd.Col = 0
                Set Grd.CellPicture = Chkno.Picture
                Grd.Col = 1
                Grd.Text = Trim(strInvNum)
                Grd.Col = 2
                Grd.Text = Trim(strPONumber)

                Grd.Col = 3
                Grd.Text = Trim(strPartNum)

                Grd.Col = 4
                Grd.Text = Trim(strQty)

                Grd.Col = 5
                Grd.Text = Trim(strUnitPrice)

                Grd.Col = 6
                Grd.Text = Trim(strInvTot)
                Grd.Col = 7
                Grd.Text = Trim(strSODate)
                Grd.Col = 8
                Grd.Text = Trim(strInvDate)
                Grd.Col = 9
                Grd.Text = Trim(strContainer)
                Grd.Col = 10
                Grd.Text = Trim(strCust)
                Grd.Col = 11
                Grd.Text = Trim(strPAUnit)

                .MoveNext
            Wend
            .Close
        End With
    End If
    
   If (Grd.Rows = 1) Then
      MsgBox "Invoices not found for selected customer.", vbExclamation, Caption
   End If
    

    MouseCursor ccArrow
    Set RdoInv = Nothing
    Exit Function

DiaErr1:
    sProcName = "fillgrid"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me

End Function

Function AddEDIFieldsLength(ByVal strType As String)
    Dim iIndex As Integer
    Dim j As Integer
    Dim RdoEdi As ADODB.Recordset
    Dim strValue As String
    Dim iTotLen As Integer
    Dim iTotalItems As Integer
    Dim iNumChar As Integer
    Dim strFields As String
    Dim strFldVal As String

    On Error GoTo DiaErr1

    If (strType <> "") Then
        sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
               & "HEADER = '" & strType & "' AND IMPORTTYPE = 'INV' ORDER BY FORATORDER"

        bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
        ReDim arrValue(0 To RdoEdi.RecordCount + 1)
        ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
        If bSqlRows Then
            With RdoEdi
                iTotalItems = 0
                While Not .EOF
                    iNumChar = !NUMCHARS
                    arrValue(iTotalItems) = CStr(iNumChar)
                    arrFieldName(iTotalItems) = !FieldName
                    iTotalItems = iTotalItems + 1
                    .MoveNext
                Wend
                .Close
            End With
        End If

    End If
    Set RdoEdi = Nothing
    Exit Function
DiaErr1:
    sProcName = "AddEDIFieldsLength"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me

End Function

Function SetPartHeader(ByVal arrFieldName As Object, ByVal arrValue As Object, _
            ByRef strPartCnt As String, ByRef strPartNum As String, ByRef strPAUnit As String)

    On Error GoTo DiaErr1


    Exit Function

DiaErr1:
    sProcName = "SetPartHeader"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "Esi2000", "EsiSale", "LastPrefix", sLastPrefix
End Sub

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' make sure that you release the Hook
    Call WheelUnHook(Me.hwnd)
    FormUnload
    Set PackPSf09a = Nothing
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

Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 0
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub


Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Function GetBuyerInfo(ByVal strCust As String, _
                  ByRef strBusPartner As String, ByRef strBusDetail As String, _
                  ByRef strBuyerCode As String)

   On Error GoTo modErr1
   Dim RdoBuy As ADODB.Recordset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS, BUYERCODE FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strBusPartner = "" & Trim(!SHPREF)
            strBusDetail = "" & Trim(!SHPDETAIL)
            strBuyerCode = "" & Trim(!BUYERCODE)
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function



Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Function strConverDate(strDate As String, ByRef strDateConv As String)
   strDateConv = Format(CDate(strDate), "yymmdd")
End Function

Private Function FormatEDIString(strInput As String, iLen As Variant, strPad As String) As String
   
   If (iLen > 0) Then
      If (strPad = "0") Then
         strInput = Format(strInput, String(iLen, "0"))
      ElseIf (strPad = "@") Then
         strInput = Format(strInput, String(iLen, "@"))
      End If
   End If

   FormatEDIString = strInput
   
End Function

Private Function CreateHeader(strCust As String, strInvNum As String, strPONumber As String, _
               strSODate As String, strInvDate As String, strContainer As String, ByRef strHeader As String)
   On Error GoTo DiaErr1
      
   Dim strBlank As String
   Dim strUnit As String
   Dim strTime As String
   Dim strInvDtConv As String
   Dim strSODateConv As String
   Dim strBusPartner As String
   Dim strBusDetail As String
   Dim strBuyerCode As String
   
   strHeader = "H"
   strUnit = "EA"
   ' Get Fields Chars
   'strContainer = "8028"
   'strGrossWt = "1987"
   'strCarton = "1234"
   strBlank = ""
   'strBusPartner = "PACCAR"
   'strBusDetail = "DE"
   
   GetBuyerInfo strCust, strBusPartner, strBusDetail, strBuyerCode
   
   ' get the Field length
   AddEDIFieldsLength "H"
   
   strInvNum = FormatEDIString(strInvNum, arrValue(0), "0")
   strConverDate strInvDate, strInvDtConv
   strInvDtConv = FormatEDIString(strInvDtConv, arrValue(1), "0")
   strBusPartner = FormatEDIString(strBusPartner, arrValue(2), "@")
   strBusDetail = strBusDetail & FormatEDIString(" ", (arrValue(3) - Len(strBusDetail)), "@")
   strPONumber = strPONumber & FormatEDIString(" ", (arrValue(4) - Len(strPONumber)), "@")
   
   strConverDate strInvDate, strInvDtConv
   strInvDtConv = FormatEDIString(strInvDtConv, arrValue(5), "0")
   
   strConverDate strSODate, strSODateConv
   strSODateConv = FormatEDIString(strSODateConv, arrValue(6), "0")
   
   strContainer = FormatEDIString(strContainer, arrValue(7), "0")
   
   strHeader = strHeader & strInvNum & strInvDtConv & strBusPartner & strBusDetail _
                  & strPONumber & strInvDtConv & strSODateConv & strContainer
   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Function CreateInvDetail(strInvNum As String, strQty As String, strPAUnits As String, _
                  strUnitPrice As String, strPartNum As String, ByRef strInvDetail As String)
   On Error GoTo DiaErr1
      
   Dim strTotItem As String
   Dim strBlank As String
   Dim strBlank1 As String
   Dim strVendPartNum As String
   strBlank = ""
   strBlank1 = ""
   
   AddEDIFieldsLength "D"
   ' Get total Items
   strInvNum = FormatEDIString(CStr(strInvNum), arrValue(0), "0")
   strQty = FormatEDIString(strQty, arrValue(1), "0")
   strPAUnits = FormatEDIString(strPAUnits, arrValue(2), "@")
   strBlank = FormatEDIString(" ", arrValue(3), "@")
   strUnitPrice = FormatEDIString(Replace(strUnitPrice, ".", ""), arrValue(4), "0")
   strPartNum = strPartNum & FormatEDIString(" ", (arrValue(5) - Len(strPartNum)), "@")
   'strBlank1 = FormatEDIString(" ", arrValue(6), "@")
   strVendPartNum = Mid(strPartNum, 1, arrValue(6))
   
   strInvDetail = "D" & strInvNum & strQty & strPAUnits & strBlank & _
            strUnitPrice & strPartNum & strVendPartNum
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateCD"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function CreateSODetail(strInvNum As String, strQty As String, _
         strInvTot As String, ByRef strSODetail As String)
   On Error GoTo DiaErr1
      
   Dim strInvTot1 As String
   strInvTot1 = strInvTot
   
   AddEDIFieldsLength "S"
   ' Get total Items
   
   strInvNum = FormatEDIString(CStr(strInvNum), arrValue(0), "0")
   strQty = FormatEDIString(strQty, arrValue(1), "0")
   strInvTot = FormatEDIString(Replace(strInvTot, ".", ""), arrValue(2), "0")
   strInvTot1 = FormatEDIString(Replace(strInvTot1, ".", ""), arrValue(3), "0")
   
   strSODetail = "S" & strInvNum & strQty & strInvTot & strInvTot1
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSODetail"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

