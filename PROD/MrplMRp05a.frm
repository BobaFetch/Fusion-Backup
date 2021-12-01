VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form MrplMRp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excess Inventory"
   ClientHeight    =   3675
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optPORcp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Tag             =   "8"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optCanPk 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox optAdj 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   2760
      Width           =   255
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   10
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp05a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp05a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "MrplMRp05a.frx":0308
      PictureDn       =   "MrplMRp05a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "MrplMRp05a.frx":0594
      PictureDn       =   "MrplMRp05a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Receipt"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   10
      Left            =   3600
      TabIndex        =   20
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inventory Activity :"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Inventory Adjustment"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled MO Pick"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "MrplMRp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
' MrplMRp05a - Raw Material Finished Goods
'
' Notes:
'
' Created: 10/28/09 (nth)
' Revisions:
'
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRp05a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(value As Integer)
   SysPrinters.Show
   ShowPrinters.value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sCustomerReport As String
   Dim sType As String
   Dim b As Byte
   
   MouseCursor 13
   On Error GoTo DiaErr1
      
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomerReport As String
   Dim sType As String
   Dim b As Byte
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   
   SetMdiReportsize MDISect
   
   For b = 1 To 4
      If typ(b) = vbChecked Then
         sType = sType & CStr(b) & ","
      End If
   Next
   If Len(sType) Then
      sType = Left(sType, Len(sType) - 1)
   End If
   
   sCustomReport = GetCustomReport("finRMFGa.rpt")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   MDISect.Crw.Formulas(2) = "AsOf='" & txtDte & "'"
   MDISect.Crw.Formulas(3) = "Title1='As Of " & txtDte & "'"
   MDISect.Crw.Formulas(4) = "Title2='Includes Part Types " & sType & "'"
   MDISect.Crw.Formulas(5) = "Title3='For Part Class " & cmbCls & "'"
   
   MDISect.Crw.Formulas(6) = "Dsc=" & optCanPk
   MDISect.Crw.Formulas(7) = "Ext=" & optAdj
   MDISect.Crw.Formulas(8) = "QOH=" & optPORcp
   
   sSql = "{InvaTable.INADATE}<=cdate('" & txtDte & _
          "') AND {PartTable.PALEVEL} IN [" & sType & "]"
   If UCase(cmbCls) <> "ALL" Then
      sSql = sSql & " AND {PartTable.PACLASS}='" & cmbCls & "'"
   End If
   
   MDISect.Crw.SelectionFormula = sSql
   
   SetCrystalAction Me
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCombo()
   FillProductClasses Me
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   
   ' Save by Menu Option
   sOptions = RTrim(typ(1).value) _
              & RTrim(typ(2).value) _
              & RTrim(typ(3).value) _
              & RTrim(typ(4).value) _
              & RTrim(optCanPk.value) _
              & RTrim(optAdj.value) _
              & RTrim(optPORcp.value)
   SaveSetting "Esi2000", "EsiProd", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "Prdmr04Printer", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      typ(1).value = Val(Mid(sOptions, 1, 1))
      typ(2).value = Val(Mid(sOptions, 2, 1))
      typ(3).value = Val(Mid(sOptions, 3, 1))
      typ(4).value = Val(Mid(sOptions, 4, 1))
      optCanPk.value = Val(Mid(sOptions, 5, 1))
      optAdj.value = Val(Mid(sOptions, 6, 1))
      optPORcp.value = Val(Mid(sOptions, 8, 1))
   Else
      typ(1).value = vbChecked
      typ(2).value = vbChecked
      typ(3).value = vbChecked
      typ(4).value = vbChecked
      optCanPk.value = vbUnchecked
      optAdj.value = vbUnchecked
      optPORcp.value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Prdmr04Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
   'lblPrinter = GetSetting("Esi2000", "EsiProd", Me.Name & TTSAVEPRN, lblPrinter)
   'If lblPrinter = "" Then
   '   lblPrinter = "Default Printer"
   'End If
End Sub

