VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaCLp11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Products Variance"
   ClientHeight    =   4275
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1920
      TabIndex        =   24
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbCode 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4800
      Picture         =   "diaCLp11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   7155
      Picture         =   "diaCLp11a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6600
      Picture         =   "diaCLp11a.frx":04CC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkJGL 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   8
      Top             =   3765
      Width           =   200
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   3465
      Width           =   200
   End
   Begin VB.CheckBox chkDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   200
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   3165
      Width           =   200
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaCLp11a.frx":064A
      PictureDn       =   "diaCLp11a.frx":0790
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4680
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4275
      FormDesignWidth =   7680
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts Like"
      Height          =   405
      Index           =   9
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   21
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Journal to G.L"
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   20
      Top             =   3765
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Only"
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   19
      Top             =   3465
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2565
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   3165
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "And :"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "M.O. Closed Between :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "diaCLp11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaCLp11a - Standard Products Variance
'
' Notes:
'
' Created: 08/09/2008
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************
Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

'Private Sub cmbPrt_LostFocus()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   If Trim(cmbPrt) = "" Then
'      cmbPrt = "ALL"
'   End If
'End Sub
'
Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
        ' Load the Product Code combo box
        PopulateCombo cmbCode, "PAPRODCODE", "PartTable"
        bOnLoad = False
        FillPartCombo cmbPrt
        cmbPrt = ""
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   'txtBeg = Format(txtEnd, "mm/01/yy")
   GetOptions
   bOnLoad = True
End Sub

Private Sub cmdVew_Click()
   ViewParts.Show
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaCLp11a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   On Error GoTo whoops
   
   'setmdireportsizemdisect
   
   'get custom report name if one has been defined
   sCustomReport = GetCustomReport("fincl11.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'pass formulas
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='From " & txtBeg & " Through " & txtEnd & " for Parts Matching " & cmbPrt & "'"
   MdiSect.crw.Formulas(3) = "PartCode='" & cmbCode & "'"
   MdiSect.crw.Formulas(4) = "ShowPartDesc=" & chkDsc
   MdiSect.crw.Formulas(5) = "ShowExtDesc=" & chkExt
   MdiSect.crw.Formulas(6) = "ShowSummary=" & chkSummary
   MdiSect.crw.Formulas(7) = "ShowGLTransferJournal=" & chkJGL
   
   'pass Crystal SQL if required
   sSql = ""
   MdiSect.crw.SelectionFormula = sSql
  ' setcrystalaction Me
   Exit Sub
   
whoops:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
  
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
     
   End If
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub PopulateCombo(cbo As ComboBox, sColumn As String, sTable As String)
   'populate combobox from database table values of a specific column
   
   cbo.Clear
   cbo.AddItem "<ALL>"
   
   Dim rdo As ADODB.Recordset
   sSql = "SELECT DISTINCT " & sColumn & " FROM " & sTable
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            If Trim(.Fields(0)) = "" Then
               cbo.AddItem "<BLANK>"
            Else
               cbo.AddItem Trim(.Fields(0))
            End If
            .MoveNext
         Loop
      End With
   End If
   cbo.ListIndex = 0
End Sub

