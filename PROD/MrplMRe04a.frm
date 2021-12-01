VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form MrplMRe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Release SC Status MO's"
   ClientHeight    =   10725
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10725
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "&Deselect All"
      Height          =   435
      Left            =   2100
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Create MO from MRP exception"
      Top             =   2760
      Width           =   1755
   End
   Begin VB.CheckBox chkPrintPickLists 
      Caption         =   "Include Pick Lists with printed MOs"
      Height          =   615
      Left            =   13320
      TabIndex        =   29
      Top             =   4020
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   660
      Width           =   3495
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRe04a.frx":0000
      Height          =   315
      Left            =   5280
      Picture         =   "MrplMRe04a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   660
      Width           =   350
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   1620
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CheckBox optDtPart 
      Caption         =   "Show Detail Parts"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   2820
      Width           =   255
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "&Select All"
      Height          =   435
      Left            =   240
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Create MO from MRP exception"
      Top             =   2760
      Width           =   1755
   End
   Begin VB.OptionButton optAll 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optExp 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10680
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdMO 
      Caption         =   "&Release MOs"
      Height          =   435
      Left            =   13320
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Create MO from MRP exception"
      Top             =   4740
      Width           =   1755
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe04a.frx":0684
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe04a.frx":09C6
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdMRP 
      Caption         =   "&Get SC MO's"
      Height          =   435
      Left            =   6000
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Get MRP exceptions / SC MO's"
      Top             =   2760
      Width           =   2355
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   660
      Width           =   3495
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRe04a.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1140
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1140
      Width           =   1250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   13320
      Top             =   9120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10725
      FormDesignWidth =   15240
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   7215
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   3360
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12726
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin ResizeLibCtl.ReSize ReSize2 
      Left            =   0
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10725
      FormDesignWidth =   15240
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   9960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail Parts"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   24
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Only Exception Part List"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   22
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show PartList with NO Exception"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   21
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7260
      Picture         =   "MrplMRe04a.frx":14B6
      Stretch         =   -1  'True
      Top             =   900
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7260
      Picture         =   "MrplMRe04a.frx":1840
      Stretch         =   -1  'True
      Top             =   660
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   15
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   14
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   13
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   645
      Width           =   1425
   End
   Begin VB.Menu mnuPopupCpy 
      Caption         =   "PopupCopy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "MrplMRe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

'Passed document stuff
Dim iDocEco As Integer
Dim strDocName As String
Dim strDocClass As String
Dim strDocSheet As String
Dim strDocDesc As String
Dim strDocAdcn As String
Dim sListRef As String
Dim strListRev As String
Dim UsingMouse As Boolean
Dim bGenMRP As Boolean
Dim RightClickPart As Boolean
Dim iPkRecord As Integer

Dim gstrPartRef As String
Dim giRunNo As Integer
Dim gstrDate As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
   
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
    
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
    sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub



'Private Sub cmbByr_LostFocus()
'   cmbByr = CheckLen(cmbByr, 20)
'   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
'   If Trim(cmbByr) = "" Then cmbByr = "ALL"
'
'End Sub


'Private Sub cmbCde_LostFocus()
'   cmbCde = CheckLen(cmbCde, 6)
'   If cmbCde = "" Then cmbCde = "ALL"
'
'End Sub
'

'Private Sub cmbCls_LostFocus()
'   cmbCls = CheckLen(cmbCls, 6)
'   If cmbCls = "" Then cmbCls = "ALL"
'
'End Sub
'

Private Sub cmbExport_Click()

   Dim sFileName As String
   
   If Trim(txtFilePath.Text) = "" Then
      MsgBox "Please specify Excel file name and directory.", vbExclamation
      Exit Sub
   End If
   
   Dim fldCnt As Integer
   Dim sFieldsToExport(40) As String
   
   fldCnt = 9
   AddFieldsToExport sFieldsToExport
   
   sFileName = txtFilePath.Text
   SaveAsExcelGrd sFieldsToExport, fldCnt, sFileName

End Sub

Private Sub SaveAsExcelGrd(ByRef aFieldsToExport() As String, iFieldCnt As Integer, ByVal filename)

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim iRow, fd, iList As Integer
    Dim iSheetsPerBook As Integer
    
    'Cell count, the cells we can use
    Dim iCell As Integer

    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo SaveToExcelError
    
    If xlApp Is Nothing Then Set xlApp = New Excel.Application

    iSheetsPerBook = xlApp.SheetsInNewWorkbook
    xlApp.SheetsInNewWorkbook = 1
    Set xlBook = xlApp.Workbooks.Add
    xlApp.SheetsInNewWorkbook = iSheetsPerBook
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlSheet = xlBook.Worksheets(1)
    
    'Get the field names
    For fd = 0 To iFieldCnt - 1
      
      xlSheet.Cells(1, fd + 1).Value = aFieldsToExport(fd)
      xlSheet.Cells(1, fd + 1).Interior.ColorIndex = 33
      xlSheet.Cells(1, fd + 1).Font.Bold = True
      xlSheet.Cells(1, fd + 1).BorderAround xlContinuous
    Next

   Dim strAssyPart As String
   Dim strRun As String
   Dim strQty As String
   Dim strReqDate As String
   Dim strActDate As String
   Dim strPickItem As String
   Dim strReqQty As String
   Dim strQOHQty As String
   Dim strRunTot As String
   
   iCell = 0
   iRow = 2
   ' Go throught all the record in the grid and create MO
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      'MM If Grd.CellPicture = Chkyes.Picture Then
      
      If Grd.CellPicture <> 0 Then
        Grd.Col = 1
        strAssyPart = Grd.Text
        
        Grd.Col = 2
        strRun = Grd.Text
        
        Grd.Col = 3
        strQty = Grd.Text
        
        Grd.Col = 4
        strReqDate = Grd.Text
        
        Grd.Col = 5
        strActDate = Grd.Text
      
        Grd.Col = 0
        If ((iList + 1) < Grd.Rows) Then
        
            Grd.row = iList + 1
            
            If (Grd.CellPicture = 0) Then
                ' only if we have detail
                iList = iList + 1
                Grd.row = iList
            Else
                Grd.row = iList
            End If
        End If
      
      End If
      
      ' Assym Part
      xlSheet.Cells(iRow, iCell + 1).Value = strAssyPart
      ' Run
      xlSheet.Cells(iRow, iCell + 2).Value = strRun
      ' Qty
      xlSheet.Cells(iRow, iCell + 3).Value = strQty
      ' Required Date
      xlSheet.Cells(iRow, iCell + 4).Value = strReqDate
      ' Action Date
      xlSheet.Cells(iRow, iCell + 5).Value = strActDate
      

      If (iList < Grd.Rows) Then
      
        If (Grd.CellPicture = 0) Then
            ' Pick
            Grd.Col = 1
            strPickItem = Grd.Text
            xlSheet.Cells(iRow, iCell + 6).Value = strPickItem
            
            ' Req Qty
            Grd.Col = 3
            strReqQty = Grd.Text
            xlSheet.Cells(iRow, iCell + 7).Value = strReqQty
            
            ' PAQOH Qty
            Grd.Col = 4
            strQOHQty = Grd.Text
            xlSheet.Cells(iRow, iCell + 8).Value = strQOHQty
            
            ' RunTot Qty
            Grd.Col = 5
            strRunTot = Grd.Text
            xlSheet.Cells(iRow, iCell + 9).Value = strRunTot
            
            
        End If
        
     End If
      
      xlSheet.Columns().AutoFit
      
      iRow = iRow + 1
    '  MM End If
      
   Next
    
   xlSheet.Rows.RowHeight = 20

   ' Save the Worksheet.
   If Len(Trim(filename)) > 0 Then
       If InStr(1, filename, ".") = 0 Then filename = filename + ".xlsx"
       xlBook.SaveAs filename
   End If
   
   Set xlSheet = Nothing
   xlBook.Close False
   Set xlBook = Nothing
   xlApp.Visible = True
   xlApp.DisplayAlerts = True
   xlApp.Quit
   Set xlApp = Nothing
   
   MsgBox "Successfully Exported The Data."
   
   Screen.MousePointer = vbArrow
   Exit Sub
   
SaveToExcelError:
   MsgBox Err.Description & " Row = " & str(iRow) & " Column = " & str(iCell)


End Sub



Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim I As Integer
   I = 0
   sFieldsToExport(I) = "Part Number"
   sFieldsToExport(I + 1) = "Run"
   sFieldsToExport(I + 2) = "Qty"
   sFieldsToExport(I + 3) = "RequiredDate"
   sFieldsToExport(I + 4) = "ActionDate"
   sFieldsToExport(I + 5) = "PickItems"
   sFieldsToExport(I + 6) = "ReqQty"
   sFieldsToExport(I + 7) = "PAQOH"
   sFieldsToExport(I + 8) = "RunTot"
   

End Function

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDeselectAll_Click()
   Dim iList As Long
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
          Set Grd.CellPicture = Chkno.Picture
          SelectRunStat 0
      End If
   Next

End Sub

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
        
   Dim iList As Long
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkno.Picture Then
          Set Grd.CellPicture = Chkno.Picture
          SelectRunStat 0
      End If
   Next

  txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If

End Sub

Private Sub cmdSelAll_Click()
   
   Dim iList As Long
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkno.Picture Then
          Set Grd.CellPicture = Chkyes.Picture
          SelectRunStat 0
      End If
   Next

End Sub


Private Sub mnuCopy_Click()
   If Grd.Col = 1 Then
      Clipboard.Clear
      Clipboard.SetText Grd.Text
   End If

End Sub

Private Sub Text1_GotFocus()
   Grd.Text = Text1.Text
   If Grd.Col >= Grd.Cols Then Grd.Col = 1
   ChangeCellText
End Sub

Private Sub Grd_EnterCell()  ' Assign cell value to the textbox
   If (bGenMRP = True) Then Text1.Text = Grd.Text
End Sub

Private Sub Grd_LeaveCell()
   ' Assign textbox value to Grd
   If (bGenMRP = True) And (Text1.Visible = True) Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If

End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
   
'   If Grd.Col <= Grd.Cols - 2 Then
'      Grd.Col = Grd.Col + 1
'      ChangeCellText
'   Else
'      If Grd.row + 1 < Grd.Rows Then
'        Grd.row = Grd.row + 1
'        Grd.Col = 1
'        ChangeCellText
'      End If
'   End If
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Text1.Move Grd.Left + Grd.CellLeft, _
   Grd.Top + Grd.CellTop, _
   Grd.CellWidth, Grd.CellHeight
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub cmdMO_Click()
   ReleaseMOs
   
   ' update list to exclude newly released MOs
   cmdMRP_Click
End Sub

Private Sub cmdMRP_Click()

   cmdMRP.Enabled = False
   
   Dim sParts As String
'   Dim sCode As String
'   Dim sClass As String
'   Dim sBuyer As String
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEndDate As String
   
   MouseCursor 13
   Grd.Clear
   GrdAddHeader
   
   GetMRPCreateDates sBegDate, sEndDate
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBDate = "1/1/2000"
   Else
      sBDate = Format(txtBeg, "mm/dd/yyyy")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "12/31/2024"
   Else
      sEDate = Format(txtEnd, "mm/dd/yyyy")
   End If
   
   If Trim(cmbPart) = "" Then cmbPart = "ALL"
'   If Trim(cmbCde) = "" Then cmbCde = "ALL"
'   If Trim(cmbCls) = "" Then cmbCls = "ALL"
'   If Trim(cmbByr) = "" Then cmbByr = "ALL"
   
   If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
   
'   If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
'   If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
'   If Trim(cmbByr) = "ALL" Then sBuyer = "" Else sBuyer = Trim(cmbByr)
   
      
'   sSql = "RptMRPMOQtyShortage '" & sParts & "', '" & sBDate & "','" & sEDate & "'"
'   clsADOCon.ExecuteSql sSql
   
   
   'need to know if customer is imaginetics
   Dim isImaginetics As Boolean, rdoIMAINC As ADODB.Recordset
   sSql = "select CONAME from ComnTable where CONAME like '%Imaginetics%'"
   isImaginetics = clsADOCon.GetDataSet(sSql, rdoIMAINC, ES_STATIC)
   
   Dim RdoMrpEx As ADODB.Recordset
   Dim strAssyPart As String
   Dim strPartDtRqd As String
   Dim strRunStat As String
   Dim strExp As String
   
   If (optExp = True) Then
      strExp = " GROUP BY BMASSYPART, MRP_ACTIONDATE HAVING   MIN(PAQRUNTOT) < 0"
   Else
      strExp = " GROUP BY BMASSYPART, MRP_ACTIONDATE HAVING   MIN(PAQRUNTOT) >= 0"
   End If
   
      Dim strParameterNames(2) As String
      Dim varParameterValues(2) As Variant
      Dim strStoredProcName As String
   
      strParameterNames(0) = "Parts"
      strParameterNames(1) = "StartDate"
      strParameterNames(2) = "EndDate"
      
      varParameterValues(0) = sParts
      varParameterValues(1) = sBDate
      varParameterValues(2) = sEDate
      
   If optAll Then
   ' use new stored procedure logic to show only items which can be allocated in RUNKPSTART order
      'sSql = "GetScMOs '" & sParts & "', '" & sBDate & "', '" & sEDate & "'"
      
      bSqlRows = clsADOCon.ExecuteStoredProcEx("GetScMOs", strParameterNames, _
                                    varParameterValues, True, RdoMrpEx)
   Else
'      sSql = "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM, " & vbCrLf _
'         & " RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD, Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART " & vbCrLf _
'         & " FROM RunsTable, PartTable, (SELECT DISTINCT BMASSYPART, MRP_ACTIONDATE " & vbCrLf _
'         & " FROM tempMrplPartShort " & strExp & ") as f " & vbCrLf _
'         & " WHERE PARTREF = RUNREF AND RUNREF LIKE '" & sParts & "%' AND RUNSTATUS = 'SC'" & vbCrLf _
'         & " AND RUNPKSTART BETWEEN '" & sBDate & "' AND '" & sEDate & " 23:00'" & vbCrLf _
'         & " AND RUNREF = f.BMASSYPART AND RUNPKSTART = f.MRP_ACTIONDATE " & vbCrLf _
'         & " order by RUNPKSTART"
'      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrpEx, ES_STATIC)
      bSqlRows = clsADOCon.ExecuteStoredProcEx("GetScMOsBlocked", strParameterNames, _
                                    varParameterValues, True, RdoMrpEx)
   End If
   
   Dim strPartRef As String
   Dim strRunNo As String
   Dim strRunOPCur As String
   Dim bRet As Boolean
   Dim pkPart As String
  

   'bSqlRows = clsADOCon.GetQuerySet(RdoMrpEx, ado)
   If bSqlRows Then
      With RdoMrpEx
         Do Until .EOF
            pkPart = Trim(!PKPARTREF)
            strPartRef = Trim(!MRP_PARTREF)
            strRunNo = Trim(!Runno)
            strRunOPCur = Trim(!RUNOPCUR)
            
            ' For Imaginetics, only release MOs where current op is for WC 120
            ' if optAll selected, this screen has already been done
            'If isImaginetics And Not optAll Then
            If isImaginetics Then
               bRet = CheckDocKitQue(strPartRef, strRunNo, strRunOPCur)
            Else
               bRet = True
            End If
            
          
            If (bRet = True) Then
               
               ' show MO line
               If pkPart = "" Then
               
                  Grd.Rows = Grd.Rows + 1
                  Grd.row = Grd.Rows - 1
                  
                  Grd.Col = 0
                  Set Grd.CellPicture = Chkno.Picture
                  Grd.Col = 1
                  Grd.Text = Trim(!MRP_PARTNUM)
                  Grd.Col = 2
                  Grd.Text = Trim(!Runno)
                  Grd.Col = 3
                  Grd.Text = Trim(!mrp_partqtyrqd)
                  Grd.Col = 4
                  Grd.Text = Trim(!MRP_PARTDATERQD)
                  Grd.Col = 5
                  Grd.Text = Trim(!MRP_ACTIONDATE)
                  Grd.Col = 6
                  Set Grd.CellPicture = picUnchecked.Picture
                  Grd.Col = 7
                  Set Grd.CellPicture = picUnchecked.Picture
                  Grd.Col = 8
                  Set Grd.CellPicture = picUnchecked.Picture
                  
                  strAssyPart = Compress(!MRP_PARTNUM)
                  strPartDtRqd = !MRP_PARTDATERQD     '!MRP_ACTIONDATE
                  strRunStat = !RUNSTATUS
                  
                  If (strRunStat = "SC") Then
                     Grd.Col = 7
                     Set Grd.CellPicture = picChecked.Picture
                  End If
                  
                  If (optExp = True) Then
                     strExp = " AND PAQRUNTOT < 0"
                  Else
                     strExp = ""
                  End If
               
               ' Show only if the detail part is requested.
               ElseIf (optDtPart = 1) Then
                                 
                  Grd.Rows = Grd.Rows + 1
                  Grd.row = Grd.Rows - 1
                  
                  Grd.Col = 1
                     Grd.Text = "  " & pkPart
                  Grd.Col = 3
                     Grd.Text = !mrp_partqtyrqd
                  Grd.Col = 4
                     Grd.Text = CStr(!PAQOH) & " - " & CStr(!Unpicked)
                  Grd.Col = 5
                     Grd.Text = !surplus
               End If
               
            End If
            .MoveNext
         Loop
      End With
   End If
   
   cmdMRP.Enabled = True
   
   Set RdoMrpEx = Nothing
   bGenMRP = True
   MouseCursor 0
   Exit Sub
   
End Sub

Private Function CheckDocKitQue(strPartRef As String, strRunNo As String, strRunOPCur As String)

   Dim RdoCurOP As ADODB.Recordset
   Dim strDocKitOp As String
   
   sSql = " SELECT opno FROM rnopTable WHERE opref = '" & strPartRef & "'" _
      & " AND oprun = " & Val(strRunNo) & " AND opcenter = '0120' "
   strDocKitOp = ""
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCurOP, ES_KEYSET)
   If bSqlRows Then
      With RdoCurOP
         strDocKitOp = Trim(!opNo)
         ClearResultSet RdoCurOP
      End With
   Else
      CheckDocKitQue = False
   End If
   
   If (strDocKitOp = "") Then
      CheckDocKitQue = False
   Else
      If (Val(strRunOPCur) = Val(strDocKitOp)) Then
           CheckDocKitQue = True
      Else
           CheckDocKitQue = False
      End If
   End If
    
   Set RdoCurOP = Nothing
   
   Exit Function
   
DiaErr1:
   sProcName = "CheckDocKitQue"

End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Dim iCurCol As Integer
      iCurCol = Grd.Col
      If Grd.row >= 1 Then
         If Grd.row = 0 Then Grd.row = 1
               
         Grd.Col = 0
         If (Grd.CellPicture = 0) Then
            Exit Sub
         End If
         
         Grd.Col = iCurCol
         
         If (Grd.Col = 0) Then
            If Grd.CellPicture = Chkyes.Picture Then
               Set Grd.CellPicture = Chkno.Picture
            Else
               Set Grd.CellPicture = Chkyes.Picture
            End If
            SelectRunStat iCurCol
            
         ElseIf ((Grd.Col = 6) Or (Grd.Col = 7) Or (Grd.Col = 8)) Then
            
            If Grd.CellPicture = picChecked.Picture Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Set Grd.CellPicture = picChecked.Picture
            End If
            SelectRunStat iCurCol
         ElseIf (Grd.Col = 5) Then
            UsingMouse = True
            Grd.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
         End If
      End If
   End If

End Sub

Private Sub Grd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iCurCol As Integer
   iCurCol = Grd.Col

   If Button = vbRightButton Then
      If Grd.Col = 1 Then
         PopupMenu mnuPopupCpy
      End If
   End If
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Dim iCurCol As Integer
   iCurCol = Grd.Col
   If Grd.row >= 1 Then
      If Grd.row = 0 Then Grd.row = 1
      
      Grd.Col = 0
      If (Grd.CellPicture = 0) Then
         Exit Sub
      End If
      
      Grd.Col = iCurCol
            
      If (Grd.Col = 0) Then
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
         SelectRunStat iCurCol
         
      ElseIf ((Grd.Col = 6) Or (Grd.Col = 7) Or (Grd.Col = 8)) Then
         
         If Grd.CellPicture = picChecked.Picture Then
            Set Grd.CellPicture = picUnchecked.Picture
         Else
            Set Grd.CellPicture = picChecked.Picture
         End If
         SelectRunStat iCurCol
      ElseIf (Grd.Col = 5) Then
         UsingMouse = True
         Grd.Text = Text1.Text
         Text1.Visible = True
         ChangeCellText
      End If
   
   End If
End Sub

Private Sub SelectRunStat(CurCol As Integer)
   
   Dim bPLSel As Boolean
   Dim bSCSel As Boolean
   Dim bMOSel As Boolean
   
   Grd.Col = 0
   bMOSel = IIf((Grd.CellPicture = Chkyes.Picture), True, False)
   
   If (bMOSel = False) And (CurCol = 0) Then
      Grd.Col = 6
      Set Grd.CellPicture = picUnchecked.Picture
      Grd.Col = 7
      Set Grd.CellPicture = picChecked.Picture
      Grd.Col = 8
      Set Grd.CellPicture = picUnchecked.Picture
       
      ' Uncheck both the image
      Exit Sub
   End If
   
   
   If (bMOSel = True) And (CurCol = 0) Then
      Grd.Col = 6
      Set Grd.CellPicture = picChecked.Picture
      Grd.Col = 7
      Set Grd.CellPicture = picUnchecked.Picture
      Grd.Col = 8
      Set Grd.CellPicture = picChecked.Picture
   End If
   
   Grd.Col = 6
   bPLSel = IIf((Grd.CellPicture = picChecked.Picture), True, False)
   Grd.Col = 7
   bSCSel = IIf((Grd.CellPicture = picChecked.Picture), True, False)
   
   If (bPLSel = False) Then
      Grd.Col = 7
      Set Grd.CellPicture = picChecked.Picture
      
      Grd.Col = 8
      Set Grd.CellPicture = picUnchecked.Picture
   Else
      If (CurCol = 7) Then
         Grd.Col = 6
         Set Grd.CellPicture = picUnchecked.Picture
         Grd.Col = 8
         Set Grd.CellPicture = picUnchecked.Picture
      ElseIf (CurCol = 6) Then
         Grd.Col = 7
         Set Grd.CellPicture = picUnchecked.Picture
         Grd.Col = 8
         Set Grd.CellPicture = picChecked.Picture
      End If
   End If
   
End Sub

Private Sub Form_Activate() '
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      'GetLastMrp
      'GetMRPDates
      'FillBuyers
      'GetOptions
''      cmbCde.AddItem "ALL"
'      FillProductCodes
'      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
'      cmbCls.AddItem "ALL"
'      FillProductClasses
'      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
      
      bOnLoad = 0
      RightClickPart = False
   End If
   MouseCursor 0
   
End Sub
Private Sub txtPrt_LostFocus()
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
End Sub


Private Sub Form_Load()
   FormLoad Me
   FormatControls
   'GetOptions
   ' Add headers
   GrdAddHeader
   bGenMRP = False
   bOnLoad = 1
   
End Sub

Private Sub GrdAddHeader()
     
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
   
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "PartNumber"
      .Col = 2
      .Text = "Run Number"
      .Col = 3
      .Text = "Qty (ReqQty)"
      .Col = 4
      .Text = "Reqd Date (QOH - Picks)"
      .Col = 5
      .Text = "Action Date (Surplus)"
      .Col = 6
      .Text = "PL Stat"
      .Col = 7
      .Text = "SC Stat"
      .Col = 8
      .Text = "Print"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 3250
      .ColWidth(2) = 1000
      .ColWidth(3) = 1200
      .ColWidth(4) = 2000
      .ColWidth(5) = 1800
      .ColWidth(6) = 700
      .ColWidth(7) = 700
      .ColWidth(8) = 700
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRe04a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
   
End Sub

'Private Sub SaveOptions()
'   Dim sOptions As String
'   Dim sCode As String * 6
'   Dim sClass As String * 4
'   Dim sBuyer As String * 20
'   sCode = cmbCde
'   sClass = cmbCls
'   sBuyer = cmbByr
'   SaveSetting "Esi2000", "EsiProd", "MrplMRe04a", sOptions
   
'End Sub

'Private Sub GetOptions()
'   Dim sOptions As String
'   On Error Resume Next
'   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr02", sOptions)
'   If Len(Trim(sOptions)) > 0 Then
'      cmbCde = Mid$(sOptions, 1, 6)
'      cmbCls = Mid$(sOptions, 7, 4)
'      cmbByr = Trim(Mid$(sOptions, 11, 20))
'   End If
'End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub


'Private Sub FillBuyers()
'   On Error GoTo DiaErr1
''   sSql = "SELECT DISTINCT MRP_POBUYER FROM MrplTable " _
''          & "WHERE MRP_POBUYER<>'' ORDER BY MRP_POBUYER"
'
'   sSql = "SELECT BYREF FROM BuyrTable ORDER BY BYREF"
'
'   AddComboStr cmbByr.hwnd, "ALL"
'   LoadComboBox cmbByr, -1
'   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
'   If Trim(cmbByr) = "" Then cmbByr = "ALL"
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub ReleaseMOs()

   Dim iList As Integer
   Dim strPartNum As String
   Dim strQty As String
   Dim strPartRqd As String
   Dim strActDate As String
   Dim strRunStat As String
   Dim strLevel As String
   Dim bPLChked As String
   Dim strRunNum As String
   Dim bAddedMO As Boolean
   
   On Error GoTo DiaErr1
   MouseCursor 13
   Err.Clear
    
   
   bAddedMO = False
   ' Go throught all the record in the grid and create MO
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 0
        Grd.row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            
            Grd.Col = 1
            strPartNum = Grd.Text
            Grd.Col = 2
            strRunNum = Grd.Text
            
            Grd.Col = 3
            strQty = Grd.Text
            Grd.Col = 4
            strPartRqd = Grd.Text
            gstrDate = strPartRqd
            Grd.Col = 5
            strActDate = Grd.Text
            ' Default va;ue
            strRunStat = "PL"
            bPLChked = False
            Grd.Col = 6
            If (Grd.CellPicture = picChecked.Picture) Then
               strRunStat = "PL"
               bPLChked = True
            End If
            
            Grd.Col = 7
            If (Grd.CellPicture = picChecked.Picture) Then
               strRunStat = "SC"
            End If
            
            Dim strPartRef As String
            Dim iRunNo As Integer
            Dim cPalevLab As Currency
            Dim cPalevExp As Currency
            Dim cPalevMat As Currency
            Dim cPalevOhd As Currency
            Dim cPalevHrs As Currency
            Dim strRouting As String
            
            Dim RdoPart As ADODB.Recordset
            
            sSql = "SELECT PARTNUM, PARTREF, PARUN, PALEVLABOR, PALEVEXP, PALEVMATL," & vbCrLf _
                      & "PALEVOH, PALEVHRS, PALEVEL, PAROUTING " & vbCrLf _
                     & " FROM PartTable WHERE PARTREF = '" & Compress(strPartNum) & "'"

            Debug.Print sSql
            
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoPart)
            If bSqlRows Then
               With RdoPart
                  strPartRef = Trim(!PartRef)
                  iRunNo = CInt(!PARUN) + 1
                  cPalevLab = !PALEVLABOR
                  cPalevExp = !PALEVEXP
                  cPalevMat = !PALEVMATL
                  cPalevOhd = !PALEVOH
                  cPalevHrs = !PALEVHRS
                  strLevel = Trim(!PALEVEL)
                  strRouting = Trim(!PAROUTING)
               End With
            End If
            ClearResultSet RdoPart
            Set RdoPart = Nothing

            ' get Routing information
'            Dim RdoRte As ADODB.Recordset
'
'            Dim strRoutType As String
'            Dim strRtNumber As String
'            Dim strRtDesc As String
'            Dim strRtBy As String
'            Dim strRtAppBy As String
'            Dim strRtAppDate As String
'
'            sSql = "SELECT * FROM RthdTable WHERE RTREF='" & Compress(strRouting) & "' "
'            bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
'            If bSqlRows Then
'               With RdoRte
'                  strRtNumber = "" & Trim(!RTNUM)
'                  strRtDesc = "" & Trim(!RTDESC)
'                  strRtBy = "" & Trim(!RTBY)
'                  strRtAppBy = "" & Trim(!RTAPPBY)
'                  If Not IsNull(!RTAPPDATE) Then
'                     strRtAppDate = Format$(!RTAPPDATE, "mm/dd/yy")
'                  Else
'                     strRtAppDate = ""
'                  End If
'                  ClearResultSet RdoRte
'               End With
'            Else
'               strRoutType = "RTEPART" & Trim(strLevel)
'               sSql = "SELECT " & strRoutType & " FROM ComnTable WHERE COREF=1"
'               Set RdoRte = clsADOCon.GetRecordSet(sSql)
'               'Set RdoRte = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
'               If Not RdoRte.BOF And Not RdoRte.EOF Then
'                  strRtNumber = "" & Trim(RdoRte.Fields(0))
'               Else
'                  strRtNumber = ""
'               End If
'               ClearResultSet RdoRte
'            End If
'            Set RdoRte = Nothing
            
            ' Open the transaction
            clsADOCon.BeginTrans
            clsADOCon.ADOErrNum = 0
            
            'make sure run is still in SC status (someone else may have updated it)
            Dim rs As ADODB.Recordset
            sSql = "select RUNSTATUS from RunsTable" & vbCrLf _
               & "where RUNREF = '" & strPartRef & "' and RUNNO = '" & Val(strRunNum) & "'"
            bSqlRows = clsADOCon.GetDataSet(sSql, rs)
            If bSqlRows Then
               Dim stat As String
               stat = rs!RUNSTATUS
               Set rs = Nothing
               If stat <> "SC" Then
                  clsADOCon.RollbackTrans
                  MsgBox "MO " & strPartRef & "-" & strRunNum & "is in " & stat & " status.  No action necessary."
                  GoTo nextmo
               End If
            End If
            
            If (Val(strRunNum) <> 0) Then
               sSql = "UPDATE RunsTable SET RUNSTATUS = '" & strRunStat & "', RUNPLDATE = '" & Format(Now, "mm/dd/yyyy") & "' WHERE RUNREF = '" _
                        & strPartRef & "' AND RUNNO = '" & Val(strRunNum) & "'"
               clsADOCon.ExecuteSql sSql
               gstrPartRef = strPartRef
               giRunNo = Val(strRunNum)
               
               AddPickList CCur(strQty)
               iRunNo = Val(strRunNum)
            Else
               MsgBox "Attempting to create an MO rather than release it.  This should never happen.  Please contact Key Methods"
'               gstrPartRef = strPartRef
'               giRunNo = Val(iRunNo)
'               ' Create new Runs and schedule
'               sSql = "INSERT INTO RunsTable (RUNREF,RUNNO,RUNSCHED," _
'                  & "RUNSTART, RUNPKSTART, RUNPLDATE," _
'                  & "RUNSTATUS,RUNQTY,RUNPRIORITY,RUNBUDLAB," _
'                  & "RUNBUDEXP,RUNBUDMAT,RUNBUDOH,RUNBUDHRS," _
'                  & "RUNREMAININGQTY,RUNRTNUM,RUNRTDESC,RUNRTBY,RUNRTAPPBY,RUNRTAPPDATE) " _
'                  & "VALUES('" & strPartRef & "'," _
'                  & Val(iRunNo) & ",'" _
'                  & strPartRqd & "','" _
'                  & strPartRqd & "','" _
'                   & strPartRqd & "','" _
'                  & Format(Now, "mm/dd/yyyy") & "','" _
'                  & strRunStat & "'," _
'                  & Val(strQty) & "," _
'                  & Val(0) & "," _
'                  & cPalevLab & "," _
'                  & cPalevExp & "," _
'                  & cPalevMat & "," _
'                  & cPalevOhd & "," _
'                  & cPalevHrs & "," _
'                  & Val(strQty) & ",'" _
'                  & strRtNumber & "','" _
'                  & strRtDesc & "','" _
'                  & strRtBy & "','" _
'                  & strRtAppBy & "','" _
'                  & strRtAppDate & "')"
'
'               clsADOCon.ExecuteSql sSql
'
'               sSql = "UPDATE PartTable SET PARUN=" & Val(iRunNo) & " " _
'                      & "WHERE PARTREF='" & strPartRef & "'"
'               clsADOCon.ExecuteSql sSql
'
'               ' Now add Routing/Run Op
'               CopyRouting Compress(strRtNumber), strPartRef, iRunNo
'
'               ' Now add Document list
'               CreateDocumentList strPartRef, iRunNo
'
'               AddPickList CCur(strQty)

            End If
            
            If clsADOCon.ADOErrNum <> 0 Then
               MsgBox "Couldn't Successfully Update..", _
                  vbInformation, Caption
               bAddedMO = False
               clsADOCon.RollbackTrans
            Else
               clsADOCon.CommitTrans
               Grd.Col = 8
               If ((Grd.CellPicture = picChecked.Picture) And (bPLChked = True)) Then
                  PrintReport Trim(strPartRef), Val(iRunNo)
                  If chkPrintPickLists.Value = 1 Then
                     PrintPickReport Trim(strPartRef), Val(iRunNo)
                  End If
               End If
               bAddedMO = True
            End If
        End If
nextmo:
    Next
    
    If (bAddedMO = True) Then
         MsgBox "Successfully created MO's.", _
                     vbInformation, Caption
    End If
    
    MouseCursor 0
    Exit Sub

DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "cmdUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function CopyRouting(strRouting As String, strPartRef As String, iRunNo As Integer)
   Dim RdoRte As ADODB.Recordset
   
   Dim iCurrentOp As Integer
   Dim strRoutType As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   'Delete possible duplicate keys
   sSql = "DELETE FROM RnopTable WHERE OPREF='" & strPartRef _
          & "' AND OPRUN=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
          & "OPPICKOP,OPSERVPART,OPQHRS,OPMHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT FROM " _
          & "RtopTable WHERE OPREF='" & Compress(strRouting) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_KEYSET)
   If bSqlRows Then
      With RdoRte
         Do Until .EOF
            On Error Resume Next
            If iCurrentOp = 0 Then iCurrentOp = !opNo
            strRoutType = "" & Trim(!OPCOMT)
            strRoutType = ReplaceString(strRoutType)
            sSql = "INSERT INTO RnopTable (OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
                   & "OPQHRS,OPMHRS,OPPICKOP,OPSERVPART,OPSUHRS,OPUNITHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT) " _
                   & "VALUES('" & strPartRef & "'," _
                   & Trim(CStr(iRunNo)) & "," _
                   & !opNo & ",'" _
                   & Trim(!OPSHOP) & "','" _
                   & Trim(!OPCENTER) & "'," _
                   & !OPQHRS & "," _
                   & !OPMHRS & "," _
                   & !OPPICKOP & ",'" _
                   & Trim(!OPSERVPART) & "'," _
                   & !OPSETUP & "," _
                   & !OPUNIT & "," _
                   & !OPSVCUNIT & ",'" _
                   & Trim(!OPTOOLLIST) & "','" _
                   & Trim(strRoutType) & "')"
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoRte
      End With
      sSql = "UPDATE RunsTable SET RUNOPCUR=" & iCurrentOp & " " _
             & "WHERE RUNREF='" & strPartRef & "' AND RUNNO=" _
             & Val(iRunNo) & " "
      clsADOCon.ExecuteSql sSql
      CopyRouting = 1
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Sub CreateDocumentList(strPartRef As String, iRunNo As Integer)
   Dim RdoList As ADODB.Recordset
   
   
   Dim iRow As Integer
   Dim sDocRef As String
   Dim sRev As String
   
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & strPartRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT MAX(DLSREV) FROM DlstTable WHERE DLSREF='" & strPartRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_KEYSET)
   If bSqlRows Then
      With RdoList
         If Not IsNull(.Fields(0)) Then
            strListRev = "" & Trim(.Fields(0))
         Else
            On Error Resume Next
            'Dummy Row for joins
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
                   & "VALUES(1,'" & strPartRef & "'," & Val(iRunNo) & ")"
            clsADOCon.ExecuteSql sSql
            Exit Sub
         End If
         ClearResultSet RdoList
      End With
   End If
   
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & strPartRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   ' In partTable the Rev is NONE, but the DocList table has a empty string
   ' 3/7/2010
   If (Trim(strListRev) = "NONE") Then
     strListRev = ""
   End If
   
   sSql = "SELECT * FROM DlstTable WHERE DLSREF='" & strPartRef & "' " _
          & "AND DLSREV='" & strListRev & "' ORDER BY DLSDOCCLASS,DLSDOCREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_KEYSET)
   If bSqlRows Then
      With RdoList
         On Error Resume Next
         Do Until .EOF
            iRow = iRow + 1
            sDocRef = GetDocInformation("" & Trim(!DLSDOCREF), "" & Trim(!DLSDOCREV))
            sProcName = "CreateDocumentList"
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF," _
                   & "RUNDLSRUNNO,RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
                   & "RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
                   & "RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
                   & "RUNDLSDOCREFECO) VALUES(" & iRow & ",'" & Compress(strPartRef) & "'," _
                   & Val(iRunNo) & ",'" & strListRev & "','" & Trim(!DLSDOCREF) & "','" _
                   & Trim(!DLSDOCREV) & "','" & strDocName & "','" & strDocDesc & "','" _
                   & strDocSheet & "','" & strDocClass & "','" & strDocAdcn & "'," _
                   & iDocEco & ")"
                   
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoList
      End With
      MouseCursor 0
   Else
      'Dummy Row for joins - Corrected 1/30/07
      On Error Resume Next
      sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
             & "VALUES(1,'" & Compress(strPartRef) & "'," & Val(iRunNo) & ")"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoList = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdocumentli"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRevisions(strPartRef As String) As String
   
   Dim RdoLst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(strPartRef) & "' ORDER BY BMHREVDATE DESC"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         GetRevisions = "" & Trim(!BMHREV)
      End With
      
      ClearResultSet RdoLst
      Set RdoLst = Nothing
   Else
      GetRevisions = ""
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "getrevisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function AddPickList(cRunqty As Currency) As Boolean
   Dim RdoLst As ADODB.Recordset
   
   Dim bGoodHeader As Byte
   Dim bGoodPl As Byte
   Dim bResponse As Byte
   Dim bOrphanedParts As Byte
   
   Dim iRow As Integer
   Dim iTotalItems As Integer
   Dim n As Integer
   
   Dim cQuantity As Currency
   Dim cConversion As Currency
   Dim cSetup As Currency
   'Dim cRunqty As Currency
   Dim sMsg As String
   Dim sBomRev As String
   
   On Error GoTo DiaErr2
   
   iPkRecord = 0
   sBomRev = GetRevisions(gstrPartRef)
      
   'determine whether any part list for this part and rev
   sSql = "SELECT BMASSYPART FROM BmplTable " & vbCrLf _
          & "WHERE BMASSYPART = '" & gstrPartRef & "'" & vbCrLf _
          & "AND BMREV = '" & sBomRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      bGoodPl = True
      ClearResultSet RdoLst
      Set RdoLst = Nothing
   Else
      MouseCursor 0
      bGoodPl = False
      MsgBox "This part does not have a parts list rev " & sBomRev, vbInformation, Caption
      AddPickList = False
      Exit Function
   End If

   sSql = "SELECT BMHREF,BMHREV,BMHOBSOLETE,BMHRELEASED,BMHEFFECTIVE " & vbCrLf _
          & "FROM BmhdTable" & vbCrLf _
          & "WHERE BMHREF='" & gstrPartRef & "' AND BMHREV='" & sBomRev & "' " & vbCrLf _
          & "AND (BMHOBSOLETE IS NULL OR BMHOBSOLETE >='" & gstrDate & "') AND BMHRELEASED=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      bGoodHeader = True
   Else
      bGoodHeader = False
   End If
   
   If Not bGoodHeader Then
      'oops the header is gone, date invalid or not released?
      MouseCursor 0
      MsgBox "The Parts List Is Not Valid, Released, " & vbCr _
         & "Or Outdated For This Part.", vbInformation, Caption
      AddPickList = False
      Exit Function
   End If
      
   sSql = "SELECT PARTREF,PAUNITS, * FROM BmplTable" & vbCrLf _
      & "LEFT OUTER JOIN PartTable ON PARTREF=BMPARTREF " & vbCrLf _
      & "WHERE BMASSYPART='" & gstrPartRef & "'" & vbCrLf _
      & "AND BMREV='" & sBomRev & "'" & vbCrLf _
      & "ORDER BY BMSEQUENCE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_DYNAMIC)
   bOrphanedParts = 0
   If bSqlRows Then
      With RdoLst
         Do Until .EOF
            If Not IsNull(!BMSETUP) Then
               cSetup = !BMSETUP
            Else
               cSetup = 0
            End If
            
            If (SetupQtyEnabled = True) Then
               cQuantity = Format(((cRunqty + cSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
            Else
               cQuantity = Format(((cRunqty * (!BMQTYREQD + !BMADDER)) + cSetup), "######0.000")
            End If
            If !BMCONVERSION <> 0 Then
               cQuantity = cQuantity / !BMCONVERSION
            End If
   
            'if phantom item, then explode it
            If !BMPHANTOM = 1 Then
               InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cQuantity
            Else
               iPkRecord = iPkRecord + 1
               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                      & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
                      & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
                      & Compress(gstrPartRef) & "'," & CStr(giRunNo) & ",9,'" & gstrDate _
                      & "'," & cQuantity & "," & cQuantity & "," & iPkRecord & "," _
                      & "'" & Trim(!PAUNITS) & "','" & Trim(!BMCOMT) & "') "
               If Len(Trim(!PartRef)) = 0 Then
                   bOrphanedParts = 1
               Else
                   clsADOCon.ExecuteSql sSql
               End If
               
            End If
               
            .MoveNext
         Loop
         ClearResultSet RdoLst
         Set RdoLst = Nothing
         AddPickList = True
      End With
         
      If bOrphanedParts Then
        MsgBox "Pick List Added Successfully. However, your BOM Parts List has Orphaned Parts." & vbCrLf & "Please Contact Fusion Support"
      End If
   Else
      MouseCursor 0
      MsgBox "Couldn't Find Items For Revision " & sBomRev & ".", vbInformation, Caption
      AddPickList = False
   End If
   Set RdoLst = Nothing
   Exit Function
   
DiaErr2:
   Set RdoLst = Nothing
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub InsertPhantom(AssyPart As String, AssyRev As String, AssyQuantity As Currency)
   Dim RdoPhn As ADODB.Recordset
   Dim iList As Integer
   Dim iTotalPhantom As Integer
   Dim cPQuantity As Currency
   Dim cPConversion As Currency
   Dim cPSetup As Currency
   
   
   sSql = "SELECT * FROM BmplTable" & vbCrLf _
      & "WHERE BMASSYPART='" & AssyPart & "'" & vbCrLf _
      & "AND BMREV='" & AssyRev & "'" & vbCrLf _
      & "ORDER BY BMSEQUENCE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPhn, ES_STATIC)
   iList = -1
   If bSqlRows Then
      With RdoPhn
         Do Until .EOF
            iList = iList + 1
            If Not IsNull(!BMSETUP) Then
               cPSetup = !BMSETUP
            Else
               cPSetup = 0
            End If
            cPQuantity = Format(((AssyQuantity + cPSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
            
            If !BMCONVERSION <> 0 Then
               cPQuantity = cPQuantity / !BMCONVERSION
            End If
            
            If !BMPHANTOM = 1 Then
               InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cPQuantity
            Else
            
               iPkRecord = iPkRecord + 1
               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                  & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
                  & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
                  & gstrPartRef & "'," & CStr(giRunNo) & ",9,'" & gstrDate _
                  & "'," & cPQuantity & "," & cPQuantity & "," & iPkRecord & "," _
                  & "'" & Trim(!BMUNITS) & "','" & Trim(!BMCOMT) & "') "
               clsADOCon.ExecuteSql sSql
            
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPhn
      End With
   End If
   Set RdoPhn = Nothing
End Sub

Private Function GetDocInformation(DocumentRef As String, DocumentRev As String) As String
   Dim RdoDoc As ADODB.Recordset
   
   sProcName = "getdocinfo"
   sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR,DOECO," _
          & "DOADCN FROM DdocTable where (DOREF='" & DocumentRef & "' " _
          & "AND DOREV='" & DocumentRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         GetDocInformation = "" & Trim(!DOREF)
         strDocName = "" & Trim(!DONUM)
         strDocClass = "" & Trim(!DOCLASS)
         strDocSheet = "" & Trim(!DOSHEET)
         strDocDesc = "" & Trim(!DODESCR)
         iDocEco = !DOECO
         strDocAdcn = "" & Trim(!DOADCN)
         ClearResultSet RdoDoc
      End With
      'strDocName = CheckStrings(strDocName)
      'strDocAdcn = CheckStrings(strDocAdcn)
   Else
      strDocName = ""
      strDocClass = ""
      strDocSheet = ""
      strDocDesc = ""
      iDocEco = 0
      strDocAdcn = ""
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetDocInformation"
   
End Function

Private Sub PrintPickReport(strPartRef As String, iRunNo As Integer)

   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error GoTo Pma01Pr
   sCustomReport = GetCustomReport("prdma01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowPickComments"
    aFormulaName.Add "ShowLots"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add 0
    aFormulaValue.Add 0
    aFormulaValue.Add 1
    aFormulaValue.Add 1
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNREF} = '" & strPartRef & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(str(iRunNo)) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
Pma01Pr:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub PrintReport(strPartRef As String, iRunNo As Integer)
   MouseCursor 13
   On Error GoTo Psh01
   sProcName = "printreport"
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sSubSql As String
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "PartNumber"
   aFormulaName.Add "RunNumber"
   aFormulaName.Add "ShowOpComments"
   aFormulaName.Add "ShowOpTime"
   aFormulaName.Add "ShowSvcParts"
   aFormulaName.Add "ShowSoAllocs"
   aFormulaName.Add "ShowDocList"
   aFormulaName.Add "ShowBOM"
   aFormulaName.Add "ShowPickList"
   aFormulaName.Add "ShowMoBudget"
   aFormulaName.Add "ShowToolList"
   aFormulaName.Add "ShowServPartDoc"
   aFormulaName.Add "ShowInternalCmt"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(strPartRef)) & "'")
   aFormulaValue.Add CStr("'" & CStr(iRunNo) & "'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'1'")
   
   sCustomReport = GetCustomReport("prdsh01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNNO} = {@Run} and {PartTable.PARTREF} = {@PartNumber}"
   cCRViewer.SetReportSelectionFormula sSql
   
'   sSubSql = "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
'            & "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
'            & "{MopkTable.PKMOPART} = {?Pm-RunsTable.RUNREF} and  " _
'            & "({MopkTable.PKTYPE} = 10 OR {MopkTable.PKTYPE} = 9)"
'            ' PKTYPE=10 is picked type and PickOpenItem = 9
'   ' set the sub sql variable pass the sub report name
'   cCRViewer.SetSubRptSelFormula "custpklist.rpt", sSubSql
'
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   DoEvents
   Exit Sub
   
Psh01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1(strPartRef As String, iRunNo As Integer)

   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error GoTo Pma01Pr

   sCustomReport = GetCustomReport("prdma01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowPickComments"
    aFormulaName.Add "ShowLots"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add 1
    aFormulaValue.Add 0
    aFormulaValue.Add 1
    aFormulaValue.Add 1
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNREF} = '" & strPartRef & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(str(iRunNo)) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
Pma01Pr:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function


