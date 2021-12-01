VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EDIAsnOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Shipping Schedule"
   ClientHeight    =   10155
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   15045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
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
      TabIndex        =   36
      ToolTipText     =   " Select All"
      Top             =   5520
      Width           =   1920
   End
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select import"
      Top             =   4440
      Width           =   4695
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   5640
      TabIndex        =   26
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdASNInfo 
         Caption         =   "Add ASN Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   " Add ASN number to PS"
         Top             =   3240
         Width           =   2160
      End
      Begin VB.TextBox txtLoadNum 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtGrossWt 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox txtCarrierNum 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   2310
         Width           =   1215
      End
      Begin VB.TextBox txtCarton 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox txtASN 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Number"
         Height          =   285
         Index           =   9
         Left            =   360
         TabIndex        =   34
         Top             =   2760
         Width           =   1185
      End
      Begin VB.Label lblLastAsn 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "Last Sales Order Entered"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last ASN"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Weight"
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   30
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carrier Ref Number"
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Number"
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   28
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Ship Num"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdGetPS 
      Caption         =   "Get PS detail"
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
   Begin VB.TextBox txtPsl 
      Height          =   285
      Left            =   9780
      MaxLength       =   8
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "New Pack Slip Number (6 char max)"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox OptSoXml 
      Caption         =   "FromXMLSO"
      Height          =   195
      Left            =   7080
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   8280
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdASN 
      Caption         =   "Create ASN  file"
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
      Left            =   6840
      TabIndex        =   12
      ToolTipText     =   " Create PS from Sales Order"
      Top             =   4200
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
      TabIndex        =   13
      ToolTipText     =   " Clear the selection"
      Top             =   4800
      Width           =   1920
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
      FormDesignHeight=   10155
      FormDesignWidth =   15045
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   15
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
      Height          =   4935
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4800
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ASN File Name"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   35
      Top             =   4440
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   33
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select PS Date"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   25
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   24
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblPrefix 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9480
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "** Part Not found in Fusion"
      Height          =   255
      Left            =   12960
      TabIndex        =   21
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "* Not Delivery Dates"
      Height          =   255
      Index           =   0
      Left            =   12960
      TabIndex        =   20
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7680
      Picture         =   "EDIAsnOut.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7680
      Picture         =   "EDIAsnOut.frx":038A
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "EDIAsnOut"
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


Private Sub ClearArrays(iSize As Integer)
    Erase arrValue
    ReDim arrValue(0 To 96, iSize)
End Sub



Private Sub cmdASNInfo_Click()
   
   Dim strGrossWt As String
   Dim strCarrierNum As String
   Dim strPS As String
   
   Dim strCarton As String
   Dim strContainer As String
   Dim strLoadNum As String
   Dim strShipNum As String
   
   Dim strPONumber As String
   Dim strPartNum As String
   Dim strQty As String
   Dim strPullNum As String
   Dim strBinNum As String
   Dim strCust As String
   
   If (Trim(txtASN.Text) = "") Then
      MsgBox "ASN number is empty. Please select a customer to get the next ASN number.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   strCust = CStr(cmbCst)
   strContainer = FormatEDIString(txtASN.Text, 6, "0")
   strShipNum = CStr(Val(strContainer))
   strCarton = txtCarton.Text
   strGrossWt = txtGrossWt.Text
   strCarrierNum = txtCarrierNum.Text
   strLoadNum = txtLoadNum.Text

   RdoCon.BeginTrans
   
   Dim iList As Integer
   Dim iTotCnt As Integer
   iTotCnt = 0
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
                  
         Grd.Col = 1
         strPS = Trim(Grd.Text)
         
         If (strGrossWt = "") Then
            strGrossWt = "0.00"
         End If
         
         If (strCarton = "") Then
            strCarton = "0"
         End If
         
'         If (strCarrierNum = "") Then
'            strCarrierNum = "Null"
'         End If
'
   
         sSql = "UPDATE PshdTable SET PSCONTAINER = '" & strContainer & "'," _
                  & " PSSHIPNO = " & strShipNum & "," _
                  & " PSLOADNO = '" & strLoadNum & "', PSGROSSLBS = " & strGrossWt & ", " _
                     & " PSCARTON = " & strCarton & ", PSCARRIERNUM = '" & strCarrierNum & "'" _
            & " FROM PshdTable " _
            & " WHERE PSNUMBER = '" & strPS & "'" _
                  & " AND PshdTable.PSINVOICE = 0"
                  '& " AND PshdTable.PSPRINTED IS NULL"
                  '& " AND PshdTable.PSSHIPPRINT = 0" _

         RdoCon.Execute sSql, rdExecDirect
         
         sSql = "UPDATE ASNInfoTable SET LASTASNNUM = '" & strShipNum & "' " _
                  & " WHERE CUREF = '" & strCust & "' AND TRUCKPLANT = 1"
                  
         RdoCon.Execute sSql, rdExecDirect
         
         Grd.Col = 2
         Grd.Text = Trim(strContainer)
         
         Grd.Col = 3
         Grd.Text = IIf(Trim(strCarton) = "Null", "", Trim(strCarton))

         Grd.Col = 7
         Grd.Text = Trim(strLoadNum)
         
         iTotCnt = iTotCnt + 1
      End If
   Next
   
   If RdoCon.RowsAffected > 0 Then
      'RdoCon.RollbackTrans
      MsgBox "Updated ASN information.", _
         vbInformation, Caption
   End If
   RdoCon.CommitTrans

End Sub

Private Sub cmdCan_Click()
   'sLastPrefix = cmbPre
   Unload Me

End Sub

Private Sub cmdGetPS_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   On Error GoTo DiaErr1
   FillGrid
   
   'Grd.Row = 1
   'Grd.Col = 2
   'txtASN.Text = Trim(Grd.Text)
   'Grd.Col = 3
   'txtCarton.Text = Trim(Grd.Text)
   'Grd.Col = 7
   'txtLoadNum.Text = Trim(Grd.Text)
   
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

Private Sub cmdASN_Click()

   Dim strFileName As String
   Dim strDate As String
   Dim strCust As String
   
   strFileName = txtEdiFilePath.Text

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
         GenerateASNFile nFileNum, strDate, strCust
      End If
      ' Close the file
      Close nFileNum
   Else
      MsgBox "Select filename to create ASN export file.", _
         vbInformation, Caption
   End If

End Sub


Private Function GetPartComm(ByVal strGetPart As String, _
            ByRef strPartNum As String, ByRef bComm As Boolean) As Byte
   Dim RdoPrt As rdoResultset
   
   On Error GoTo DiaErr1
   bComm = False
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = GetDataSet(RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            strPartNum = "" & Trim(!PARTNUM)
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
   Dim RdoCst As rdoResultset
   sCust = Compress(strCusName)
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & strCusName & "'"
   bSqlRows = GetDataSet(RdoCst)
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

   Dim RdoCus As rdoResultset
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF FROM CustTable WHERE CUNAME = '" & strCusFullName & "'"
   bSqlRows = GetDataSet(RdoCus)
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
   
   Dim strMaxASN  As String
   
   FindCustomer Me, cmbCst, False
   
   If bOnLoad = 0 Then
      ' Filter the records if selected.
      strMaxASN = GetLastASN(CStr(cmbCst))
      
      If (Trim(strMaxASN) <> "") Then
         txtASN = Val(strMaxASN) + 1
      Else
         txtASN = ""
         
      End If
   End If
   
End Sub

Private Sub cmbCst_LostFocus()
'   cmbCst = CheckLen(cmbCst, 10)
'   FindCustomer Me, cmbCst, False
'   lblNotice.Visible = False
   
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
   If bOnLoad = 1 Then
   
      'FillCustomers
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
               & " A.CUREF = b.CUREF AND TRUCKPLANT = 1"
               
      LoadComboBox cmbCst, -1
      AddComboStr cmbCst.hwnd, "" & Trim("ALL")
      cmbCst = "ALL"
      txtNme = "*** All Customer selected ***"
      
      'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
   
      Dim ps As New ClassPackSlip
      Dim strMaxASN As String
      lblPrefix = ps.GetPackSlipPrefix
      txtPsl = ""
      txtPsl.MaxLength = 8 - Len(lblPrefix)
      strMaxASN = GetLastASN(CStr(cmbCst))
      
      If (Trim(strMaxASN) <> "") Then
         txtASN = Val(strMaxASN) + 1
      End If
      txtDte = Format(ES_SYSDATE, "mm/dd/yy")

      txtCarton.Text = ""
      txtGrossWt.Text = ""
      txtCarrierNum.Text = ""
      txtLoadNum.Text = ""

      'GetPackslip True
      bOnLoad = 0
   End If
   MouseCursor (0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiSale", "LastPrefix", sLastPrefix
   
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
      .ColAlignment(10) = 1
   
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "PackSlip"
      .Col = 2
      .Text = "ASN Num"
      .Col = 3
      .Text = "Carton Num"
      .Col = 4
      .Text = "PO Number"
      .Col = 5
      .Text = "PartNumber"
      .Col = 6
      .Text = "Qty"
      .Col = 7
      .Text = "Load Num"
      .Col = 8
      .Text = "Via"
      .Col = 9
      .Text = "Pull Num"
      .Col = 10
      .Text = "Bin Num"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1200
      .ColWidth(5) = 2000
      .ColWidth(6) = 700
      .ColWidth(7) = 1000
      .ColWidth(8) = 1500
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
    bOnLoad = 1

End Sub

Function GenerateASNFile(nFileNum As Integer, strDate As String, strCust As String) As Integer
   
   Dim rdoPS As rdoResultset
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim strFileName As String
   Dim strPONumber As String
   Dim strPartNum As String
   Dim strPiPartRef As String
   Dim strPSNum As String
   Dim strQty As String
   Dim strCarton As String
   Dim strContainer As String
   Dim strPrevContainer As String
   Dim strLoadNum As String
   Dim strPSVia As String
   Dim bPartFound As Boolean
   Dim bIncRow As Boolean
   Dim strPullNum As String
   Dim strBinNum As String
   Dim strShipNo As String
   Dim strGrossWt As String
   Dim strCarrierNum As String
   Dim iItem As Integer
   Dim bSelected As Boolean
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSSHIPNO, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
            & "ISNULL(PSGROSSLBS, '0.00') PSGROSSLBS,ISNULL(PSCARRIERNUM, '') PSCARRIERNUM, " _
            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, PARTNUM, ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM " _
         & " From PshdTable, psitTable, sohdTable, SoitTable, Parttable " _
         & " WHERE PshdTable.PSDATE = '" & strDate & "'" _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PARTREF = PIPART" _
          & " AND PshdTable.PSCUST IN " _
          & " (SELECT DISTINCT a.CUREF " _
          & "     FROM ASNInfoTable a, custtable b WHERE " _
          & "        A.CUREF = b.CUREF AND TRUCKPLANT = 1)" _
          & " ORDER BY PSSHIPNO"
          
          'PshdTable.PSCUST = '" & strCust & "' AND
          '& " AND PSCONTAINER <> ''"
          '& " AND PshdTable.PSINVOICE = 0 " _

   Debug.Print sSql
   
   bSqlRows = GetDataSet(rdoPS, rdOpenStatic)
   
   strPrevContainer = ""
   bSelected = False
   If bSqlRows Then
      
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strContainer = Trim(!PSCONTAINER)
         strShipNo = Trim(!PSSHIPNO)
         strPONumber = Trim(!PsNumber)
         strPartNum = Trim(!PARTNUM)
         strPiPartRef = Trim(!PIPART)
         strCarton = Trim(!PSCARTON)
         strLoadNum = Trim(!PSLOADNO)
         strPSVia = Trim(!PSVIA)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strPullNum = Trim(!PULLNUM)
         strBinNum = Trim(!BINNUM)
         strGrossWt = Trim(!PSGROSSLBS)
         strCarrierNum = Trim(!PSCARRIERNUM)
         
         If (CheckSelected(strPSNum, strPONumber, strPiPartRef)) Then
            
            If ((strPrevContainer = "") Or (strPrevContainer <> strContainer)) Then
               
               Dim strBusPartner As String
               Dim strBusDetail As String
               Dim strBuyerCode As String
               GetBuyerInfo strCust, strBusPartner, strBusDetail, strBuyerCode
                              
               ' Add Header detail
               Dim strHeader As String
               CreateHeader strCust, strContainer, strCarton, strGrossWt, _
                     strLoadNum, strCarrierNum, strBusPartner, strBusDetail, strHeader
               
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeader
                  Debug.Print strHeader
               End If
               
               ' Add CD
               Dim strHeadCD As String
               Dim iTotItems As Integer
               iTotItems = TotalPsSelected(strContainer)
               
               CreateCD strContainer, strLoadNum, iTotItems, strHeadCD
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeadCD
                  Debug.Print strHeadCD
               End If
               
               ' Add H2
               Dim strHeader2 As String
               CreateHeader2 strContainer, strPSVia, strBuyerCode, strHeader2
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeader2
                  Debug.Print strHeader2
               End If
               
               
               ' Add
               Dim strHeadR1 As String
               CreateR1 strContainer, strCarrierNum, strHeadR1
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeadR1
                  Debug.Print strHeadR1
               End If
               
               ' Add Shipping info
               Dim strN1 As String
               Dim strN2 As String
               CreateShipInfo strCust, strContainer, strBuyerCode, strN1, strN2
               
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strN1
                  Print #nFileNum, strN2
                  Debug.Print strN1
                  Debug.Print strN2
               End If
               
               strPrevContainer = strContainer
            End If
            
            ' Not add the Details
            Dim strDT As String
            CreateDetail strContainer, strPartNum, strQty, strPONumber, _
               strPSNum, strPullNum, strDT
            
            If EOF(nFileNum) Then
               Print #nFileNum, strDT
               Debug.Print strDT
            End If
            ' If any selcted set the dirty flag to true
            bSelected = True
            
         End If
         
         .MoveNext
      Wend
      .Close
      End With
   
      If (bSelected = True) Then
         MsgBox "ASN File created.", vbExclamation, Caption
      End If
      
   End If

   MouseCursor ccArrow
   Set rdoPS = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GenerateASNFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Function FillGrid() As Integer
   
   Dim strSenderCode As String
   Dim rdoPS As rdoResultset
   
   MouseCursor ccHourglass
   Grd.Rows = 1
   On Error GoTo DiaErr1
       
       
   Dim strCust, strPONumber, strPartNum As String
   Dim strPSNum As String
   Dim strDate, strQty As String
   Dim strCarton, strContainer As String
   Dim strLoadNo, strPSVia As String
   Dim bPartFound, bIncRow As Boolean
   Dim strPullNum, strBinNum As String
   Dim iItem As Integer

   strDate = txtDte.Text
   strCust = cmbCst.Text
   
   If (Trim(strCust) = "ALL") Then
      strCust = ""
   End If
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM " _
         & " From PshdTable, psitTable, sohdTable, SoitTable " _
         & " WHERE PshdTable.PSDATE = '" & strDate & "'" _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PshdTable.PSCUST IN " _
          & " (SELECT DISTINCT a.CUREF " _
          & "         FROM ASNInfoTable a, custtable b WHERE " _
          & "     A.CUREF = b.CUREF AND TRUCKPLANT = 1)"
         
           '& PshdTable.PSCUST LIKE '" & strCust & "%' AND
          '& " AND PshdTable.PSINVOICE = 0"
          '& " AND PshdTable.PSPRINTED IS NULL"
          '& " AND PshdTable.PSSHIPPRINT = 0" _

   Debug.Print sSql
   
   bSqlRows = GetDataSet(rdoPS, rdOpenStatic)
   
   If bSqlRows Then
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strContainer = Trim(!PSCONTAINER)
         strPONumber = Trim(!PsNumber)
         strPartNum = Trim(!PIPART)
         strCarton = Trim(!PSCARTON)
         strLoadNo = Trim(!PSLOADNO)
         strPSVia = Trim(!PSVIA)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strPullNum = Trim(!PULLNUM)
         strBinNum = Trim(!BINNUM)
         
         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         bIncRow = False
         iItem = 1
         
         Grd.Col = 0
         Set Grd.CellPicture = Chkno.Picture
         Grd.Col = 1
         Grd.Text = Trim(strPSNum)
         Grd.Col = 2
         Grd.Text = Trim(strContainer)
         
         Grd.Col = 3
         Grd.Text = Trim(strCarton)
         
         Grd.Col = 4
         Grd.Text = Trim(strPONumber)
         
         Grd.Col = 5
         Grd.Text = Trim(strPartNum)
         
         Grd.Col = 6
         Grd.Text = Trim(strQty)
         Grd.Col = 7
         Grd.Text = Trim(strLoadNo)
         Grd.Col = 8
         Grd.Text = Trim(strPSVia)
         Grd.Col = 9
         Grd.Text = Trim(strPullNum)
         Grd.Col = 10
         Grd.Text = Trim(strBinNum)
         
         .MoveNext
      Wend
      .Close
      End With
   End If

   MouseCursor ccArrow
   Set rdoPS = Nothing
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
   Dim RdoEdi As rdoResultset
   Dim strValue As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   
   On Error GoTo DiaErr1
   
   If (strType <> "") Then
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'ASN' ORDER BY FORATORDER"
      
      bSqlRows = GetDataSet(RdoEdi, rdOpenStatic)
      ReDim arrValue(0 To RdoEdi.RowCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RowCount + 1)
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

   Exit Function
DiaErr1:
   sProcName = "DecodeEdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Function SetPartHeader(ByVal arrFieldName As Variant, ByVal arrValue As Variant, _
            ByRef strPartCnt As String, ByRef strPartNum As String, ByRef strPAUnit As String)

   On Error GoTo DiaErr1
      
   
   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'FormUnload
    Set EDIAsnOut = Nothing
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

Private Sub CmdClear_Click()
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


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    
   Dim RdoCon As rdoResultset
   
   On Error GoTo ERR1
      
   bSqlRows = GetDataSet(RdoCon, ES_FORWARD)
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

Private Function GetLastASN(strCst As String) As String
   
   Dim rdoPS As rdoResultset
   Dim strASN As String
   strASN = ""
   
   If (Trim(strCst) = "" Or strCst = "ALL") Then
      strASN = ""
      lblLastAsn = ""
   Else
      ' Get the ship VIA information
      sSql = "SELECT MAX(PSSHIPNO) as MAXASN FROM PshdTable,ASNInfoTable" _
               & " WHERE  PSCUST = CUREF AND PSCUST = '" & strCst & "' AND TRUCKPLANT = 1" _
            & " GROUP BY PSCUST"
   
      bSqlRows = GetDataSet(rdoPS, ES_FORWARD)
      If bSqlRows Then
         With rdoPS
            lblLastAsn = "" & Trim(!MAXASN)
            strASN = "" & Trim(!MAXASN)
            ClearResultSet rdoPS
         End With
      End If
   End If
   
   GetLastASN = strASN
End Function

Private Function GetBuyerInfo(ByVal strCust As String, _
                  ByRef strBusPartner As String, ByRef strBusDetail As String, _
                  ByRef strBuyerCode As String)

   On Error GoTo ModErr1
   Dim RdoBuy As rdoResultset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS, BUYERCODE FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = GetDataSet(RdoBuy, ES_FORWARD)
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
   
ModErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


Private Function GetShipInfo(ByVal strCust As String, ByRef strFrmVnd As String, _
                  ByRef strFrmVndID As String, ByRef strFromAddrs As String, _
                  ByRef strToVnd As String, ByRef strToVndID As String, ByRef strToAddrs As String)

   On Error GoTo ModErr1
   Dim RdoBuy As rdoResultset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = GetDataSet(RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strFrmVnd = "" & Trim(!SHPFRMIDCODE)
            strFrmVndID = "" & Trim(!SHPFRMCODEQUAL)
            strFromAddrs = "U.S. CASTINGS LLC."
            strToVnd = "" & Trim(!SHPTOIDCODE)
            strToVndID = "" & Trim(!SHPTOCODEQUAL)
            strToAddrs = "" & Trim(!SHPADDRS)
            
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
ModErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


Private Sub GetPackslip(bFillText As Boolean)
   Dim ps As New ClassPackSlip
   'lblLastAsn = ps.GetLastPackSlipNumber

   If bFillText Then
      txtPsl = Right(ps.GetNextPackSlipNumber, txtPsl.MaxLength)
   End If
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Function strConverDate(strDate As String, ByRef strDateConv As String)
   strDateConv = Format(CDate(strDate), "yymmdd")
End Function

Private Function TotalPsSelected(strContainer As String) As Integer
   Dim iList As Integer
   Dim iTotCnt As Integer
   Dim strTCont As String
   
   iTotCnt = 0
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         Grd.Col = 2
         strTCont = Grd.Text
         If (strTCont = strContainer) Then
            iTotCnt = iTotCnt + 1
         End If
      End If
   Next

   TotalPsSelected = iTotCnt
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
   

Private Function CheckSelected(strPSNum As String, _
         strPONumber As String, strPartNum As String) As Boolean

   Dim bChecked As Boolean
   Dim strGrdPS As String
   Dim strGrdPO As String
   Dim strGrdPN As String
   Dim iList As Integer
   
   bChecked = False
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 1
         strGrdPS = Trim(Grd.Text)
         Grd.Col = 4
         strGrdPO = Trim(Grd.Text)
         Grd.Col = 5
         strGrdPN = Trim(Grd.Text)
               
         If ((strGrdPS = strPSNum) And (strGrdPO = strPONumber) And _
            (strGrdPN = strPartNum)) Then
               
            bChecked = True
            Exit For
         End If
      End If
   Next
   
   CheckSelected = bChecked
End Function

Private Function CreateHeader(strCust As String, strContainer As String, strCarton As String, _
               strGrossWt As String, strLoadNum As String, strCarrierNum As String, _
               strBusPartner As String, strBusDetail As String, _
               ByRef strHeader As String)
   On Error GoTo DiaErr1
      
   Dim strHeader1 As String
   Dim strBlank As String
   Dim strUnit As String
   Dim strTime As String
   Dim strDateConv As String
   
   strHeader = "H1"
   strUnit = "LB"
   ' Get Fields Chars
   'strContainer = "8028"
   'strGrossWt = "1987"
   'strCarton = "1234"
   strBlank = ""
   'strBusPartner = "PACCAR"
   'strBusDetail = "CH"
   
   ' get the Field lenght
   AddEDIFieldsLength "H1"
   
   strContainer = FormatEDIString(strContainer, arrValue(0), "0")
   strBusPartner = FormatEDIString(strBusPartner, arrValue(1), "@")
   strBusDetail = strBusDetail & FormatEDIString(" ", (arrValue(2) - Len(strBusDetail)), "@")
   strHeader = strHeader & strContainer & strBusPartner & strBusDetail
   
   strConverDate txtDte, strDateConv
   
   strDateConv = FormatEDIString(strDateConv, arrValue(3), "0")
   strTime = "170000"
   strTime = FormatEDIString(strTime, arrValue(4), "0")
   strGrossWt = FormatEDIString(strGrossWt, arrValue(5), "0")
   strUnit = FormatEDIString(strUnit, arrValue(6), "0")
   
   strHeader = strHeader & strDateConv & strTime & strGrossWt & strUnit
   
   
   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function CreateCD(strContainer As String, strLoadNum As String, iTotItems As Integer, _
                  ByRef strHeadCD As String)
   On Error GoTo DiaErr1
      
   Dim strTotItem As String
   
   AddEDIFieldsLength "CD"
   ' Get total Items
   strLoadNum = FormatEDIString(CStr(strLoadNum), arrValue(1), "@")
   strTotItem = FormatEDIString(CStr(iTotItems), arrValue(2), "0")
   
   strHeadCD = "CD" & strContainer & strLoadNum & strTotItem
   
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateCD"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Function CreateHeader2(strContainer As String, strPSVia As String, _
         strBuyerCode As String, ByRef strHeader2 As String)
   On Error GoTo DiaErr1
      
   Dim strTransMethod As String
   Dim strEquipDesc As String
   
   AddEDIFieldsLength "H2"
   ' Get total Items
   strTransMethod = "M"
   strEquipDesc = "TL"
   
   strPSVia = strPSVia & FormatEDIString(" ", (arrValue(2) - Len(strPSVia)), "@")
   strTransMethod = strTransMethod & FormatEDIString(" ", (arrValue(3) - Len(strTransMethod)), "@")
   strEquipDesc = strEquipDesc & FormatEDIString(" ", (arrValue(4) - Len(strTransMethod)), "@")
   strHeader2 = "H2" & strContainer & strBuyerCode & strPSVia & strTransMethod & strEquipDesc
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateHeader2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function CreateR1(strContainer As String, strCarrierNum As String, _
                  ByRef strHeadR1 As String)
   On Error GoTo DiaErr1
      
   Dim strCarrType As String
   Dim strBuyerCode As String
   
   AddEDIFieldsLength "R1"
   
   strCarrType = FormatEDIString("CN", arrValue(1), "@")
   strCarrierNum = FormatEDIString(strCarrierNum, arrValue(2), "0")
   strHeadR1 = "R1" & strContainer & strCarrType & strCarrierNum
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateHeader2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function CreateShipInfo(strCust As String, strContainer As String, strBuyerCode As String, _
               ByRef strN1 As String, ByRef strN2 As String)
   On Error GoTo DiaErr1
      
   Dim strFromAddrs As String
   Dim strFrmVndID As String
   Dim strFrmVnd As String
   Dim strShpFrom As String
   Dim strToAddrs As String
   Dim strToVndID As String
   Dim strToVnd As String
   Dim strShpTo As String
   
      
   GetShipInfo strCust, strFrmVnd, strFrmVndID, strFromAddrs, _
               strToVnd, strToVndID, strToAddrs
      
   AddEDIFieldsLength "N1"
   ' Get total Items
   strFromAddrs = strFromAddrs & FormatEDIString(" ", (arrValue(2) - Len(strFromAddrs)), "@")
   strFrmVndID = FormatEDIString(strFrmVndID, arrValue(3), "0")
   'strFrmVnd = Format(strFrmVnd, String(arrValue(4), "0"))
   strN1 = "N1" & strContainer & strBuyerCode & "SF" & strFromAddrs & strFrmVndID & strFrmVnd
   
   AddEDIFieldsLength "N1"
   ' Get total Items
   strToAddrs = strToAddrs & FormatEDIString(" ", (arrValue(2) - Len(strToAddrs)), "@")
   strToVndID = FormatEDIString(strToVndID, arrValue(3), "0")
   strN2 = "N1" & strContainer & strBuyerCode & "ST" & strToAddrs & strToVndID & strToVnd
   
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateShipInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Function CreateDetail(strContainer As String, strPartNum As String, _
               strQty As String, strPONumber As String, strPS As String, _
                  strPullNum As String, ByRef strDT As String)
   On Error GoTo DiaErr1
      
         Dim strVendPartNum As String
         
         AddEDIFieldsLength "DT"
         ' Get total Items
         strVendPartNum = Mid(strPartNum, 1, arrValue(2))
         strPartNum = strPartNum & FormatEDIString(" ", (arrValue(1) - Len(strPartNum)), "@")
         strQty = FormatEDIString(strQty, arrValue(3), "0")
         strPONumber = strPONumber & FormatEDIString(" ", (arrValue(4) - Len(strPONumber)), "@")
         strPS = FormatEDIString(Mid(strPS, 3, (Len(strPS) - 2)), arrValue(5), "0")
         strPullNum = strPullNum ' Shows as 9 characters 'MM & FormatEDIString(strPullNum, arrValue(6), "0")
   
         strDT = "DT" & strContainer & strPartNum & strVendPartNum & strQty & strPONumber & strPS & strPullNum

   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Public Sub ShowCalendar(frm As Form, Optional iAdjust As Integer)

   'display date selection calendar
   
   Dim iAdder As Integer
   Dim lLeft As Long
   Dim lTop As Long
   Dim sDate As Date
   
   Dim combo As Control
   Set combo = frm.ActiveControl
   
   
   If IsDate(frm.ActiveControl.Text) Then
      'frm.ActiveControl.AddItem frm.ActiveControl.Text
   Else
      'frm.ActiveControl.AddItem Format(Now, "mm/dd/yy")
      combo.Text = Format(Now, "mm/dd/yy")
   End If
   
   'set form to pass date back to
   Set SysCalendar.FromForm = frm

   'On Error Resume Next
   'See if there is a date in the combo
   If IsDate(frm.ActiveControl.Text) Then
      sDate = Format(frm.ActiveControl.Text, "mm/dd/yy")
   Else
      sDate = Format(ES_SYSDATE, "mm/dd/yy")
   End If
   
   'MM SysCalendar.Move lLeft, lTop

   ' MM bSysCalendar = True
   If IsDate(sDate) Then SysCalendar.Calendar1.value = Format(sDate, "mm/dd/yyyy")
   
   'if parent form is modal, we must show this as modal too
   On Error Resume Next
   SysCalendar.Calendar1.Refresh
   DoEvents
   SysCalendar.Show
   If Err Then
      SysCalendar.Show vbModal
   End If
   'refresh it so that it doesn't blink out
   SysCalendar.Calendar1.Refresh
   'combo.Refresh
   
End Sub


