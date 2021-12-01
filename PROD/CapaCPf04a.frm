VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form CapaCPf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Op Completed by Work Center"
   ClientHeight    =   3255
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   2160
      Width           =   4335
   End
   Begin VB.ComboBox cboShop 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   660
      Width           =   1815
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cboWorkCenter 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter New (12 Char) Or Select From List"
      Top             =   1020
      Width           =   1815
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1380
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1380
      Width           =   1250
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "CapaCPf04a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPf04a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3255
      FormDesignWidth =   7155
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6480
      Top             =   2880
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
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   11
      Left            =   300
      TabIndex        =   14
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations Scheduled Date From"
      Height          =   465
      Index           =   9
      Left            =   300
      TabIndex        =   12
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   8
      Left            =   3180
      TabIndex        =   11
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5655
      TabIndex        =   10
      Top             =   1380
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1770
      TabIndex        =   8
      Top             =   1980
      Width           =   105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   1020
      Width           =   1695
   End
End
Attribute VB_Name = "CapaCPf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtDte, 4)
   
End Sub

Private Sub GetOptions()
   'Get By Menu Option
   On Error Resume Next
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub cboShop_Click()
   FillWorkCenters
End Sub

Private Sub cboShop_LostFocus()
   FillWorkCenters
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If

End Sub

Private Sub cmbExport_Click()

   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
   ExportOpCompetedForWC
   
   
End Sub

Private Function ExportOpCompetedForWC()

   Dim sParts As String
   Dim sCenter As String
   Dim sShop As String
   
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sFileName As String
   
   On Error GoTo ExportError

   Dim RdoPO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(10) As String
   AddFieldsToExport sFieldsToExport
   
   sCenter = Compress(cboWorkCenter)
   If sCenter = "ALL" Then sCenter = ""
   
   sShop = cboShop.Text
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtDte) = "" Then txtDte = "ALL"
   If Not IsDate(txtBeg) Then
      sBDate = "01/01/2000"
   Else
      sBDate = Format(txtBeg, "mm/dd/yyyy")
   End If
   If Not IsDate(txtDte) Then
      sEDate = "12/31/2024"
   Else
      sEDate = Format(txtDte, "mm/dd/yyyy")
   End If

    sSql = "select OPREF, OPRUN, OPNO, OPSHOP, OPCENTER, OPCOMT, " & vbCrLf
    sSql = sSql & "OPSCHEDDATE , opcompdate, OPYIELD, OPACCEPT " & vbCrLf
    sSql = sSql & "from rnopTable where opcomplete =1 AND " & vbCrLf
    sSql = sSql & "   opcompdate Between '" & sBDate & "' AND '" & sEDate & "'" & vbCrLf
    sSql = sSql & "   AND OPSHOP LIKE '" & sShop & "%' AND OPCENTER LIKE '" & sCenter & "%'"
    

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      SaveAsExcel RdoPO, sFieldsToExport, sFileName
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If

   Set RdoPO = Nothing
   Exit Function
   
ExportError:
   MouseCursor 0
   cmbExport.Enabled = True
   MsgBox Err.Description
   

End Function

Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "OPREF"
   sFieldsToExport(i + 1) = "OPRUN"
   sFieldsToExport(i + 2) = "OPNO"
   sFieldsToExport(i + 3) = "OPSHOP"
   sFieldsToExport(i + 4) = "OPCENTER"
   sFieldsToExport(i + 5) = "OPCOMT"
   sFieldsToExport(i + 6) = "OPSCHEDDATE"
   sFieldsToExport(i + 7) = "opcompdate"
   sFieldsToExport(i + 8) = "OPYIELD"
   sFieldsToExport(i + 9) = "OPACCEPT"

End Function

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad <> 0 Then
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   cboWorkCenter = ""
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPf04a = Nothing
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Trim(txtBeg) = "" Then
      txtBeg = "ALL"
   Else
      txtBeg = CheckDateEx(txtBeg)
   End If
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   If Trim(txtDte) = "" Then
      txtDte = "ALL"
   Else
      txtDte = CheckDateEx(txtDte)
   End If
   
End Sub



Private Sub cboWorkCenter_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cboWorkCenter_LostFocus()
   cboWorkCenter = CheckLen(cboWorkCenter, 12)
   If Len(cboWorkCenter) = 0 Then cboWorkCenter = "ALL"
   
End Sub

Private Sub FillShops()
   Dim wc As New ClassWorkCenter
   wc.PopulateShopCombo cboShop, cboWorkCenter
End Sub

Private Sub FillWorkCenters()
   Dim wc As New ClassWorkCenter
   wc.PoulateWorkCenterCombo cboShop, cboWorkCenter
End Sub


