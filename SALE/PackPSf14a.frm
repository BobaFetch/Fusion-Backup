VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PackPSf14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get ASN Information"
   ClientHeight    =   9270
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   16770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   16770
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select import"
      Top             =   9480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGetPS 
      Caption         =   "Get ASN detail"
      Height          =   360
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2145
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1320
      Width           =   1555
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
      Left            =   7560
      TabIndex        =   6
      ToolTipText     =   " Create PS from Sales Order"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf14a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
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
      FormDesignHeight=   9270
      FormDesignWidth =   16770
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
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   2280
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   12091
      _Version        =   393216
      Rows            =   3
      Cols            =   14
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
      Left            =   960
      TabIndex        =   12
      Top             =   9480
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select PS Date"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   3375
   End
End
Attribute VB_Name = "PackPSf14a"
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

Dim strPartNum As String

Private txtKeyPress As New EsiKeyBd



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
   
   Exit Sub
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmbCst_Click()
   
   Dim strMaxASN  As String
   
   FindCustomer Me, cmbCst, False
End Sub

Private Sub cmbCst_LostFocus()
'   cmbCst = CheckLen(cmbCst, 10)
'   FindCustomer Me, cmbCst, False
'   lblNotice.Visible = False
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
   
      'FillCustomers
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
               & " A.CUREF = b.CUREF AND TRUCKPLANT = 1"
               
      LoadComboBox cmbCst, -1
      AddComboStr cmbCst.hWnd, "" & Trim("ALL")
      cmbCst = "ALL"
      txtNme = "*** All Customer selected ***"
      
      'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
      
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
      .Text = "PackSlip"
      .Col = 1
      .Text = "Container"
      .Col = 2
      .Text = "Cust"
      .Col = 3
      .Text = "ASN Num"
      .Col = 4
      .Text = "Carton Num"
      .Col = 5
      .Text = "Gross LBS"
      .Col = 6
      .Text = "PartNumber"
      .Col = 7
      .Text = "Carrier Num"
      .Col = 8
      .Text = "Load Num"
      .Col = 9
      .Text = "Via"
      .Col = 10
      .Text = "PO Number"
      .Col = 11
      .Text = "Qty"
      .Col = 12
      .Text = "Pull Num"
      .Col = 13
      .Text = "Bin Num"
      

      .ColWidth(0) = 1000
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1000
      .ColWidth(6) = 1700
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 1200
      .ColWidth(11) = 1000
      .ColWidth(12) = 1000
      .ColWidth(13) = 1200
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub


Function FillGrid() As Integer
   
   Dim strSenderCode As String
   Dim rdoPS As ADODB.Recordset
   
   MouseCursor ccHourglass
   Grd.Rows = 1
   On Error GoTo DiaErr1
       
       
   Dim strCust, strPONumber, strPartNum As String
   Dim strPSNum As String
   Dim strDate, strQty As String
   Dim strCarton, strContainer As String
   Dim strGrossLbs As String
   Dim strShipNo  As String
   
   Dim strCarrierNum As String
   Dim strLoadNo, strPSVia As String
   Dim bIncRow As Boolean
   Dim strPullNum, strBinNum As String
   Dim iItem As Integer

   strDate = txtDte.Text
   strCust = cmbCst.Text
   
   If (Trim(strCust) = "ALL") Then
      strCust = ""
   End If
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSCUST, PSSHIPNO, ISNULL(PSCARTON, '') PSCARTON," _
            & " ISNULL(PSGROSSLBS, '0.00') PSGROSSLBS,ISNULL(PSCARRIERNUM, '') PSCARRIERNUM," _
            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, PARTNUM, ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM" _
            & " From PshdTable, psitTable, sohdTable, SoitTable, Parttable" _
         & " WHERE PshdTable.PSDATE = '" & strDate & "'" _
          & " AND PshdTable.PSCUST LIKE '" & strCust & "%'" _
            & " AND PSNUMBER = PIPACKSLIP" _
            & " AND SONUMBER = ITSO" _
            & " AND ITPSNUMBER = ITPSNUMBER" _
            & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
            & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
            & " AND SoitTable.ITREV = PsitTable.PISOREV" _
            & " AND PARTREF = PIPART" _
            & " AND PshdTable.PSCUST IN" _
                  & " (SELECT DISTINCT a.CUREF" _
                  & " FROM ASNInfoTable a, custtable b WHERE" _
                  & " A.CUREF = b.CUREF AND TRUCKPLANT = 1)" _
                  & " ORDER BY PSSHIPNO"
   
   
'   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
'            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM " _
'         & " From PshdTable, psitTable, sohdTable, SoitTable " _
'         & " WHERE PshdTable.PSDATE = '" & strDate & "'" _
'          & " AND PshdTable.PSCUST LIKE '" & strCust & "%'" _
'          & " AND PSNUMBER = PIPACKSLIP" _
'          & " AND SONUMBER = ITSO" _
'          & " AND ITPSNUMBER = ITPSNUMBER" _
'          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
'          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
'          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
'          & " AND PshdTable.PSCUST IN (SELECT DISTINCT a.CUREF " _
'          & "                FROM ASNInfoTable a, custtable b WHERE " _
'          & "                A.CUREF = b.CUREF AND TRUCKPLANT = 1)"
          
          '" & strCust & "%'
          ' MM & " AND PshdTable.PSINVOICE = 0"
          '& " AND PshdTable.PSPRINTED IS NULL"
          '& " AND PshdTable.PSSHIPPRINT = 0" _

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, adOpenStatic)
   
   If bSqlRows Then
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strContainer = Trim(!PSCONTAINER)
         strCust = Trim(!PSCUST)
         strShipNo = Trim(!PSSHIPNO)
         strCarton = Trim(!PSCARTON)
         strGrossLbs = Trim(!PSGROSSLBS)
         strPartNum = Trim(!PARTNUM)
         strCarrierNum = Trim(!PSCARRIERNUM)
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
         Grd.Text = Trim(strPSNum)
         Grd.Col = 1
         Grd.Text = Trim(strContainer)
         Grd.Col = 2
         Grd.Text = Trim(strCust)
         Grd.Col = 3
         Grd.Text = Trim(strShipNo)
         
         Grd.Col = 4
         Grd.Text = Trim(strCarton)
         
         Grd.Col = 5
         Grd.Text = Trim(strGrossLbs)
         
         Grd.Col = 6
         Grd.Text = Trim(strCarrierNum)
         Grd.Col = 7
         Grd.Text = Trim(strLoadNo)
         Grd.Col = 8
         Grd.Text = Trim(strPSVia)
         Grd.Col = 9
         Grd.Text = Trim(strPONumber)
         
         Grd.Col = 10
         Grd.Text = Trim(strPartNum)
         Grd.Col = 11
         Grd.Text = Trim(strQty)
         Grd.Col = 12
         Grd.Text = Trim(strPullNum)
         Grd.Col = 13
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


Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'FormUnload
    Set PackPSf14a = Nothing
End Sub


Private Function strConverDate(strDate As String, ByRef strDateConv As String)
   strDateConv = Format(CDate(strDate), "yyyymmdd")
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
   

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
End Sub


