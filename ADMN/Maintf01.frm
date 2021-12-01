VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form Maintf01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance: Correct Part Quantities"
   ClientHeight    =   2370
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "Analyze"
      Height          =   435
      Left            =   2880
      TabIndex        =   12
      Top             =   1680
      Width           =   1515
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "Part quantity versus inventory activity"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "Part quantity versus lot header quantity"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   3735
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "Adjust lot item qtys to sum to LOTREMAINGQTY"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   4155
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Maintf01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "Maintf01.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "Maintf01.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2370
      FormDesignWidth =   7260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "     Errors     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   11
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblErrors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4980
      TabIndex        =   10
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label lblErrors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4980
      TabIndex        =   8
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblErrors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   4980
      TabIndex        =   6
      Top             =   360
      Width           =   915
   End
End
Attribute VB_Name = "Maintf01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdAnalyze_Click()

   Dim i As Integer
   Dim view As String
   Dim rdo As rdoResultset
   
   For i = 0 To 2
      Me.lblErrors(i) = ""
      If Me.chkSelect(i).Value = vbChecked Then
         view = GetView(i)
         If view <> "" Then
            sSql = "select count(*) from " & view
            If GetDataSet(rdo, ES_FORWARD) Then
               Me.lblErrors(i) = rdo.rdoColumns(0)
            End If
         End If
      End If
   Next

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



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub GetOptions()

End Sub

Private Sub SaveOptions()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
End Sub

Private Sub PrintReport()
   
   'display counts
   cmdAnalyze_Click
   
   Dim tbl As String
   tbl = "MaintQtyHealth"
   MouseCursor ccHourglass
   
   'drop table if it exists
   On Error Resume Next
   sSql = "drop table " & tbl
   RdoCon.Execute sSql
   On Error GoTo DiaErr1
   
   'create table
   sSql = "create table " & tbl & vbCrLf _
      & "(" & vbCrLf _
      & "   QtyCol varchar(30)," & vbCrLf _
      & "   SumCol varchar(30)," & vbCrLf _
      & "   PartRef varchar(30)," & vbCrLf _
      & "   LotNumber varchar(15)," & vbCrLf _
      & "   Qty decimal(16,4)," & vbCrLf _
      & "   SumQty decimal(16,4)" & vbCrLf _
      & ")"
   RdoCon.Execute sSql
   
   'populate table with selected statistics
   Dim i As Integer
   Dim view As String
   Dim rdo As rdoResultset
   
   For i = 0 To 2
      If Me.chkSelect(i).Value = vbChecked Then
         view = GetView(i)
         If view <> "" Then
            sSql = "insert into " & tbl & vbCrLf _
               & "select * from " & view
            RdoCon.Execute sSql
         End If
      End If
   Next

   'now create report
   SetMdiReportsize MDISect
   sCustomReport = GetCustomReport("Maint01")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "RequestedBy='Requested By: " & sInitials & "'"
   sSql = ""
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me

   MouseCursor ccDefault
   Exit Sub
   
DiaErr1:
   MouseCursor ccDefault
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Function GetView(ViewNo As Integer)
   Select Case ViewNo
   Case 0
      GetView = "viewMaintLotRemQtyVsLoiQty"
   Case 1
      GetView = "viewMaintPaQohVsLotRemQty"
   Case 2
      GetView = "viewMaintPaQohVsInaQty"
   Case Else
      MsgBox "Unknown view number " & ViewNo

   End Select
End Function

