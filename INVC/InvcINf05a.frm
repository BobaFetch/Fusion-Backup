VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InvcINf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Part Code and Class from Excel"
   ClientHeight    =   8025
   ClientLeft      =   1845
   ClientTop       =   1080
   ClientWidth     =   12660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClass 
      Caption         =   "Update Prod Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10320
      TabIndex        =   10
      ToolTipText     =   " Create Cash Receipts"
      Top             =   3360
      Width           =   1920
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8040
      TabIndex        =   9
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   2160
      Width           =   1920
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "Update Prod Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10320
      TabIndex        =   8
      ToolTipText     =   " Create Cash Receipts"
      Top             =   2760
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
      Height          =   480
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   2160
      Width           =   1920
   End
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Excel data"
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
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
      FormDesignHeight=   8025
      FormDesignWidth =   12660
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   8880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4935
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7320
      Picture         =   "InvcINf05a.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7320
      Picture         =   "InvcINf05a.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1305
   End
End
Attribute VB_Name = "InvcINf05a"
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
Dim bUnLoad As Boolean

Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sSJARAcct As String
Dim sCrCommAcct As String
Dim sCrRevAcct As String
Dim sCrExpAcct As String
Dim sTransFeeAcct As String
Dim strOffsetAcct As String

Dim sAccount As String
Dim sMsg As String

Private txtKeyPress As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me

End Sub

Private Sub cmdClass_Click()
   Dim strMsg As String
   
   strMsg = "Class"
   If (strMsg <> "") Then
      UpdatePartDetail strMsg
   End If

End Sub

Private Sub cmdCode_Click()
   
   Dim strMsg As String
   strMsg = "Code"
   
   If (strMsg <> "") Then
      UpdatePartDetail strMsg
   End If
End Sub

Private Function UpdatePartDetail(strMsg As String)
   Dim iList As Long
   Dim strPartNum As String
   Dim sSql As String
   Dim strFieldVal As String
   Dim strField As String
   
   On Error GoTo DiaErr1
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         Grd.Col = 1
         strPartNum = Grd.Text
         If strPartNum <> "" Then
            
            
            If (strMsg = "Code") Then
               Grd.Col = 3
               strFieldVal = Grd.Text
               strField = "PAPRODCODE"
            Else
               Grd.Col = 5
               strFieldVal = Grd.Text
               strField = "PACLASS"
            End If
            
            sSql = "UPDATE PartTable SET " & strField & " = '" & strFieldVal & "' WHERE PARTNUM = '" & strPartNum & "'"
            clsADOCon.ExecuteSQL sSql '
            
         End If
      End If
   Next
      
   If (clsADOCon.ADOErrNum = 0) Then
      MsgBox ("Updated selected Product " & strMsg)
   Else
      MsgBox ("There was a Error, while updated the Product " & strMsg)
   End If
   
   Exit Function
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

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
   Dim strFilePath As String
   
   On Error GoTo DiaErr1
   strFilePath = txtXLFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select a Excel file to create Cash Receipt.", _
            vbInformation, Caption
      Exit Sub
   End If

   MouseCursor 13
   DeleteOldData ("ImpPartCdCls")
   ParsePartDetail (strFilePath)
   

   sSql = "SELECT PARTNUM, PAPRODCODE, PACLASS, IMPRODCODE, IMPRODCLS " _
            & "  FROM PArtTable, ImpPartCdCls WHERE PARTNUM = IMPARTNUM AND" _
            & "(IMPRODCODE <> PAPRODCODE OR IMPRODCLS <> PACLASS)" _

   FillGrid (sSql)
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Excel Files (*.xls) | *.xls|"
   
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtXLFilePath.Text = ""
   Else
       txtXLFilePath.Text = fileDlg.FileName
   End If
End Sub

Private Sub cmdSel_Click()
   Dim iList As Integer
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      Set Grd.CellPicture = Chkyes.Picture
   Next
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      bOnLoad = 0
   End If
    
   MouseCursor (0)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   'Call WheelUnHook(Me.hWnd)
   
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
   
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Cur ProdCode"
      .Col = 3
      .Text = "New ProdCode"
      .Col = 4
      .Text = "Cur ProdClass"
      .Col = 5
      .Text = "New ProdClass"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 2500
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   

   'Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Function FillGrid(sSql As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim iItem  As Integer
   Dim strPartNum As String
   Dim strPartType As String
   Dim strPCode As String
   Dim strPCls As String
   Dim strImpPCode As String
   Dim strImpPCls As String
   

   Debug.Print sSql
   
   Dim RdoExo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
   
   Grd.Rows = 1
   If bSqlRows Then
      With RdoExo
      While Not .EOF
         strPartNum = !PartNum
         strPCode = !PAPRODCODE
         strPCls = !PACLASS
         strImpPCode = !IMPRODCODE
         strImpPCls = !IMPRODCLS
         
         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         
         Grd.Col = 0
         Set Grd.CellPicture = Chkno.Picture
            
         Grd.Col = 1
         Grd.Text = Trim(strPartNum)
         Grd.Col = 2
         Grd.Text = Trim(strPCode)
         Grd.Col = 3
         Grd.Text = Trim(strImpPCode)
         Grd.Col = 4
         Grd.Text = Trim(strPCls)
         Grd.Col = 5
         Grd.Text = Trim(strImpPCls)
         
         .MoveNext
         
      Wend
      .Close
      End With
   End If
   
   Set RdoExo = Nothing
   MouseCursor ccArrow
       
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
    FormUnload
    Set InvcINf05a = Nothing
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


Private Function DeleteOldData(strTableName As String)

   If (strTableName <> "") Then
      sSql = "DELETE FROM " & strTableName
      clsADOCon.ExecuteSQL sSql
   End If

End Function

Private Function ParsePartDetail(strFullPath As String)

   Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim strPartNum As String
   Dim strPartType As String
   Dim strPartCode As String
   Dim strPartClass As String
   Dim iIndex As Integer
   Dim bContinue As Boolean
   
   On Error GoTo DiaErr1
   
   If (strFullPath <> "") Then
      Set xlApp = New Excel.Application
   
      Set wb = xlApp.Workbooks.Open(strFullPath)
   
      Set ws = wb.Worksheets(1) 'Specify your worksheet name
      
      bContinue = True
      iIndex = 2
      While (bContinue)
      
         strPartNum = ws.Cells(iIndex, 1)
         strPartType = ws.Cells(iIndex, 2)
         strPartCode = ws.Cells(iIndex, 3)
         strPartClass = ws.Cells(iIndex, 4)
         
         If (strPartNum <> "") Then
   
            sSql = "INSERT INTO ImpPartCdCls (IMPARTNUM, IMPARTTYPE, IMPRODCODE,IMPRODCLS) " _
               & "VALUES('" & strPartNum & "','" _
                     & strPartType & "','" & strPartCode & "','" _
                     & strPartClass & "')"
            Debug.Print sSql
            
            clsADOCon.ExecuteSQL sSql
         
         End If
         
         
         If (strPartNum = "") Then
            bContinue = False
         End If
         
         strPartNum = ""
         strPartType = ""
         strPartCode = ""
         strPartClass = ""
         iIndex = iIndex + 1
      Wend
      
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
   End If
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
