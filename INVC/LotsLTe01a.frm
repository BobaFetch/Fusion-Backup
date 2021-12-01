VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form LotsLTe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Lots"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtCert 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "User Lot (40)"
      Top             =   3000
      Width           =   3615
   End
   Begin VB.ComboBox lblPrinter 
      Height          =   315
      Left            =   240
      TabIndex        =   73
      Top             =   6600
      Width           =   3975
   End
   Begin VB.CommandButton optPrn 
      Height          =   320
      Left            =   4250
      Picture         =   "LotsLTe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Print The Report"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   490
   End
   Begin VB.CommandButton optDisLbl 
      Height          =   320
      Left            =   4800
      Picture         =   "LotsLTe01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Display The Report"
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   490
   End
   Begin VB.TextBox txtTotLabor 
      Height          =   285
      Left            =   2700
      TabIndex        =   9
      Tag             =   "1"
      Top             =   4320
      Width           =   900
   End
   Begin VB.TextBox txtTotOh 
      Height          =   285
      Left            =   4740
      TabIndex        =   11
      Tag             =   "1"
      Top             =   4320
      Width           =   900
   End
   Begin VB.TextBox txtTotExp 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4320
      Width           =   900
   End
   Begin VB.TextBox txtTotMatl 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Tag             =   "1"
      Top             =   4320
      Width           =   900
   End
   Begin VB.ComboBox cboExpirationDate 
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTe01a.frx":0308
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "LotsLTe01a.frx":0AB6
      DownPicture     =   "LotsLTe01a.frx":1428
      Height          =   350
      Left            =   5520
      Picture         =   "LotsLTe01a.frx":1D9A
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Standard Comments"
      Top             =   2280
      Width           =   350
   End
   Begin VB.TextBox txtSplt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Splits Only 20 Char Alpha/Numeric"
      Top             =   3360
      Width           =   2150
   End
   Begin VB.TextBox lblNumber 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "System Produced Lot Number Click To Set User Lot The Same"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Change The User Lot ID"
      Top             =   1080
      Width           =   875
   End
   Begin VB.ComboBox cmbLot 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select User Lot Number From The List"
      Top             =   1080
      Width           =   4005
   End
   Begin VB.TextBox txtCmt 
      Height          =   675
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Comments (2048)"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton optDis 
      Height          =   350
      Left            =   6000
      Picture         =   "LotsLTe01a.frx":239C
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Print or View Detail"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.TextBox txtUnitCost 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Tag             =   "1"
      Top             =   3720
      Width           =   900
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Tag             =   "3"
      ToolTipText     =   "Storage Location For This Lot"
      Top             =   4680
      Width           =   675
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5700
      TabIndex        =   25
      Tag             =   "1"
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2220
      TabIndex        =   24
      Tag             =   "3"
      ToolTipText     =   "Select Type From List (Or Blank)"
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Left            =   5340
      TabIndex        =   23
      Tag             =   "1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Left            =   3660
      TabIndex        =   22
      Tag             =   "1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optSal 
      Caption         =   "SO"
      Height          =   255
      Left            =   3420
      TabIndex        =   21
      ToolTipText     =   "Select On Allocation Type"
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optMal 
      Caption         =   "MO"
      Height          =   255
      Left            =   2220
      TabIndex        =   20
      ToolTipText     =   "Select On Allocation Type"
      Top             =   7320
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtWum 
      Height          =   285
      Left            =   4920
      TabIndex        =   18
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   5760
      Width           =   435
   End
   Begin VB.TextBox txtLum 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   5400
      Width           =   435
   End
   Begin VB.TextBox txtHum 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   5040
      Width           =   435
   End
   Begin VB.TextBox txtWid 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "Mat Width"
      Top             =   5760
      Width           =   915
   End
   Begin VB.TextBox txtLng 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Mat Length"
      Top             =   5400
      Width           =   915
   End
   Begin VB.TextBox txtHgt 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Mat Heght"
      Top             =   5040
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   6732
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6180
      Top             =   5280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7050
      FormDesignWidth =   7155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cert / HT #"
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   75
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Lot Label"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   27
      Left            =   4200
      TabIndex        =   74
      ToolTipText     =   "Print Lot Label"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblTotCost 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   70
      Top             =   4320
      Width           =   900
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "= Tot Cost"
      Height          =   255
      Index           =   31
      Left            =   5760
      TabIndex        =   69
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ Exp"
      Height          =   255
      Index           =   30
      Left            =   3720
      TabIndex        =   68
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matl"
      Height          =   255
      Index           =   29
      Left            =   1680
      TabIndex        =   67
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ Labor"
      Height          =   255
      Index           =   26
      Left            =   2700
      TabIndex        =   66
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ OH"
      Height          =   255
      Index           =   28
      Left            =   4740
      TabIndex        =   65
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Totals"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   64
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label lblExtCost 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   63
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "= Exended Cost"
      Height          =   255
      Index           =   24
      Left            =   4560
      TabIndex        =   62
      Top             =   3765
      Width           =   1215
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3420
      TabIndex        =   61
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "x Qty"
      Height          =   255
      Index           =   23
      Left            =   2760
      TabIndex        =   60
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label lblExpirationDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Expiration Date"
      Height          =   255
      Left            =   4080
      TabIndex        =   59
      Top             =   3420
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split Comments"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   56
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   55
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Remaining"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   52
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   51
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   50
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   49
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   48
      Top             =   3780
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   120
      TabIndex        =   47
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   46
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   45
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Temporarily Disabled)"
      Height          =   255
      Index           =   20
      Left            =   4500
      TabIndex        =   44
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   19
      Left            =   5100
      TabIndex        =   43
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Number"
      Height          =   255
      Index           =   18
      Left            =   660
      TabIndex        =   42
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   315
      Index           =   17
      Left            =   4860
      TabIndex        =   41
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3300
      TabIndex        =   40
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   16
      Left            =   660
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocate To"
      Height          =   255
      Index           =   15
      Left            =   660
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   37
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   36
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   35
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   34
      Top             =   5820
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   33
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   32
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   31
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Comments"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   1395
   End
End
Attribute VB_Name = "LotsLTe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Revised in total as requested 4/8/03
'9/1/04 omit tools
'8/10/05 Added Split Lot Comments
'8/24/05 Changed tab order and GetThisLot from LotsLTe01b
Option Explicit
Dim RdoCur As ADODB.Recordset
Dim bGoodLot As Byte
Dim bGoodPart As Byte
Dim bOnLoad As Byte
Dim bTotalLots As Byte

Dim iIndex As Integer
Dim cOldCost As Currency
Dim sOldLot As String
'Dim sLots(250, 2) As String  0 = lot number, 1 = user lot id (not used)
Dim sLots() As String   'lot numbers
Private LotsExpire As Boolean

Public quantityPerLabel As Currency
Public totalQuantity As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboExpirationDate_LostFocus()
    Dim sTempDate As String
    
   'if valid date save it
   If LotsExpire And bGoodLot = 1 Then
      
      If Not IsDate(cboExpirationDate.Text) Then
         'cboExpirationDate.SetFocus
         Exit Sub
      End If
      
      Dim daysInPast As Long
'      daysInPast = DateDiff("d", CDate(cboExpirationDate.Text), Now)
        sTempDate = Format(cboExpirationDate.Text, "mm/DD/20yy")
        daysInPast = DateDiff("d", CDate(sTempDate), Now)
        
      If daysInPast > 0 Then
         Select Case MsgBox("The expiration date is " & daysInPast & " days in the past.  Is that correct?", vbYesNo)
         Case vbYes
         Case Else
            cboExpirationDate.SetFocus
            Exit Sub
         End Select
      End If
      
      If IsDate(cboExpirationDate.Text) Then
         With RdoCur
            If IsNull(!LOTEXPIRESON) Or !LOTEXPIRESON <> CDate(cboExpirationDate.Text) Then
'Debug.Print "save exp date " & cboExpirationDate.Text
               !LOTEXPIRESON = CDate(cboExpirationDate.Text)
               .Update
            End If
         End With
      Else
      End If
   End If

End Sub

Private Sub cmbLot_Click()
   'On Error Resume Next
   If cmbLot.ListIndex = -1 Then cmbLot.ListIndex = 0
   'lblNumber = sLots(cmbLot.ListIndex, 0)
   lblNumber = sLots(cmbLot.ListIndex)
   bGoodLot = GetThisLot()
   
End Sub


Private Sub cmbLot_LostFocus()
   'On Error Resume Next
   If Trim(cmbLot) = "" And cmbLot.ListCount > 0 Then
      cmbLot = cmbLot.List(0)
   End If
   bGoodLot = GetThisLot()
   If bGoodLot Then cmdChg.Enabled = True _
                                     Else cmdChg.Enabled = False
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetLotPart(Compress(cmbPrt))
   cmdChg.Enabled = False
   
End Sub


Private Sub cmbPrt_LostFocus()
   bGoodPart = GetLotPart(Compress(cmbPrt))
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
End Sub


'Private Sub cmbPrt_Validate(Cancel As Boolean)
'   If Not CostsAreBalanced Then
'      Cancel = True
'   End If
'End Sub
'
Private Sub cmdCan_Click()
   
   If LotsExpire And bGoodLot = 1 Then
      
      If Not IsDate(cboExpirationDate.Text) Then
         Select Case MsgBox("Valid lot expiration date required.  Do you want to exit anyway?", vbYesNo)
         Case vbYes
            Unload Me
            Exit Sub
         End Select
         cboExpirationDate.SetFocus
         Exit Sub
      End If
         
      If IsNull(RdoCur!LOTEXPIRESON) Or RdoCur!LOTEXPIRESON <> CDate(cboExpirationDate.Text) Then
         
         Dim daysInPast As Long
         daysInPast = DateDiff("d", CDate(cboExpirationDate.Text), Now)
         If daysInPast > 0 Then
            Select Case MsgBox("The expiration date is " & daysInPast & " days in the past.  Is that correct?", vbYesNo)
            Case vbYes
            Case Else
               cboExpirationDate.SetFocus
               Exit Sub
            End Select
         End If
      End If
   End If
   
'   If lblExtCost.Caption <> lblTotCost.Caption Then
'      Select Case MsgBox("The extended and total costs do not agree.  Do you want to exit anyway?", vbYesNo)
'      Case vbYes
'      Case Else
'         txtUnitCost.SetFocus
'         Exit Sub
'      End Select
'   End If
'
   If CostsAreBalanced Then
      Unload Me
   End If
   
End Sub



Private Sub cmdChg_Click()
   LotsLTe01b.txtlot = cmbLot
   LotsLTe01b.Show
   
End Sub


Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 3
      SysComments.Show
      cmdComments = False
   End If
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5501"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub





Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   Dim X As Printer
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   cmbPrt.Enabled = True
   'populate printer selection combo
   For Each X In Printers
      If Left(X.DeviceName, 9) <> "Rendering" Then
            lblPrinter.AddItem X.DeviceName
      End If
   Next
   
   On Error Resume Next
   
   Dim sDefaultPrinter As String
   If lblPrinter.ListCount > 0 Then
      sDefaultPrinter = lblPrinter.List(0)
   End If
   
   lblPrinter.Text = GetSetting("Esi2000", "EsiInv", "LotLabelPrinter", sDefaultPrinter)
    
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   SaveSetting "Esi2000", "EsiInv", "LotLabelPrinter", lblPrinter.Text
   Set RdoCur = Nothing
   Set LotsLTe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblNumber.BackColor = Me.BackColor
   txtSplt.BackColor = Me.BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PATOOL,LOTPARTREF FROM " _
          & "PartTable,LohdTable WHERE (PARTREF=LOTPARTREF AND PATOOL=0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0) " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If bSqlRows Then
      If cmbPrt.ListCount > 0 Then
         cmbPrt = cmbPrt.List(0)
         'bGoodPart = GetLotPart(Compress(cmbPrt))
      Else
         lblDsc = "No Parts with Lots Found"
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub optDis_Click()
   Dim sDate As String
   Dim sVendor As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
      sCustomReport = GetCustomReport("lotdetail")
      Set cCRViewer = New EsCrystalRptViewer
      cCRViewer.Init
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
      aFormulaName.Add "CompanyName"
       
      aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
       
      cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

      sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
            
      cCRViewer.SetReportSelectionFormula (sSql)
      cCRViewer.CRViewerSize Me
      cCRViewer.ShowGroupTree False
      cCRViewer.SetDbTableConnection
   
      cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
      cCRViewer.ClearFieldCollection aFormulaName
      cCRViewer.ClearFieldCollection aFormulaValue
      Set cCRViewer = Nothing
   
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.ReportFileName = sReportPath & "lotdetail.rpt"
'   sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
'   MdiSect.Crw.SelectionFormula = sSql
'   MdiSect.Crw.Destination = crptToWindow
'   MdiSect.Crw.Action = 1
'   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub


Private Sub optDisLbl_Click()
   PrintReport
End Sub
Private Sub optMal_Click()
   If optMal.Value = True Then
      z1(16).Visible = False
      z1(17).Visible = False
      lblSon.Visible = False
      cmbSon.Visible = False
      cmbItm.Visible = False
      
      z1(18).Visible = True
      z1(19).Visible = True
      cmbMon.Visible = True
      cmbRun.Visible = True
   Else
      z1(16).Visible = True
      z1(17).Visible = True
      lblSon.Visible = True
      cmbSon.Visible = True
      cmbItm.Visible = True
      
      z1(18).Visible = False
      z1(19).Visible = False
      cmbMon.Visible = False
      cmbRun.Visible = False
   End If
   
End Sub



Private Sub optPrn_Click()
   PrintReport
End Sub
Private Function PrintReport()

   Dim strPartRef As String
   Dim strLotNum As String
   Dim strUserLot As String
   Dim strOrgQty As String
   Dim strLotLoc As String
   
   strPartRef = Compress(cmbPrt)
   strLotNum = Me.lblNumber
   strUserLot = Me.lblNumber
   strOrgQty = Me.lblRem 'Me.lblQty
   strLotLoc = Me.txtLoc
   
   If (CDbl(strOrgQty) <= 0) Then
      MsgBox "Lot remaining Qty is zero." & vbCr & "Can not print the Lot label.", _
         vbInformation, Caption
      Exit Function
   End If
   
   Load LotsLTe01c
   LotsLTe01c.lblPartNo = strPartRef
   LotsLTe01c.lblLotNum = strLotNum
   LotsLTe01c.lblUserLotNum = cmbLot.Text
   LotsLTe01c.lblLocation = strLotLoc
   LotsLTe01c.lblQty = strOrgQty
   LotsLTe01c.txtQtyPerLabel = strOrgQty
   
   totalQuantity = strOrgQty

   Set LotsLTe01c.ParentForm = Me
   LotsLTe01c.Show vbModal
   If Me.lblRem > 0 And Me.quantityPerLabel > 0 Then
      PrintLabels strPartRef, strLotNum, Me.lblRem, Me.quantityPerLabel
   End If

End Function

Private Sub PrintLabels(strPartRef As String, strLotNum As String, totalQuantity As Currency, quantityPerLabel As Currency)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   Dim quantityLeft As Currency
   quantityLeft = totalQuantity
   Do
      
      sCustomReport = GetCustomReport("LotsLTe01.rpt")
      Set cCRViewer = New EsCrystalRptViewer
      cCRViewer.Init
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
      aFormulaName.Add "Quantity"
       
      If quantityLeft > quantityPerLabel Then
         aFormulaValue.Add CStr("'" & quantityPerLabel & "'")
      Else
         aFormulaValue.Add CStr("'" & quantityLeft & "'")
      End If
       
      cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

      sSql = "{lohdtable.LOTNUMBER} = '" & strLotNum & "' AND {lohdtable.LOTPARTREF} = '" & strPartRef & "'"
            
      cCRViewer.SetReportSelectionFormula (sSql)
      cCRViewer.CRViewerSize Me
      cCRViewer.ShowGroupTree False
      cCRViewer.SetDbTableConnection
   
      cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
      cCRViewer.ClearFieldCollection aFormulaName
      cCRViewer.ClearFieldCollection aFormulaValue
      Set cCRViewer = Nothing
      
      quantityLeft = quantityLeft - quantityPerLabel
   
   Loop While quantityLeft > 0
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optSal_Click()
   If optMal.Value = True Then
      z1(16).Visible = False
      z1(17).Visible = False
      lblSon.Visible = False
      cmbSon.Visible = False
      cmbItm.Visible = False
      
      z1(18).Visible = True
      z1(19).Visible = True
      cmbMon.Visible = True
      cmbRun.Visible = True
   Else
      z1(16).Visible = True
      z1(17).Visible = True
      lblSon.Visible = True
      cmbSon.Visible = True
      cmbItm.Visible = True
      
      z1(18).Visible = False
      z1(19).Visible = False
      cmbMon.Visible = False
      cmbRun.Visible = False
   End If
   
End Sub



Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTCOMMENTS = txtCmt
         .Update
      End With
   End If
   
End Sub




Public Sub GetLots(Optional HideId As Byte)
   Dim RdoCmb As ADODB.Recordset
   Dim iRow As Integer
   On Error GoTo DiaErr1
   cmbLot.Clear
   Erase sLots
   ReDim sLots(1000)
   iRow = -1
   sSql = "SELECT DISTINCT LOTNUMBER,LOTUSERLOTID,LOTPARTREF FROM LohdTable,LoitTable WHERE " _
         & " LOTNUMBER=LOINUMBER AND LOITYPE <> 16 AND " _
          & "LOTPARTREF='" & Compress(cmbPrt) & "' ORDER BY LOTUSERLOTID"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            iRow = iRow + 1
            AddComboStr cmbLot.hWnd, "" & Trim(!LOTUSERLOTID)
            
            'if lot array full, add another 1000 elements
            If iRow > UBound(sLots, 1) Then
               ReDim Preserve sLots(iRow + 999) As String
            End If
            
'            sLots(iRow, 0) = "" & Trim(!LotNumber)
'            sLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
            sLots(iRow) = "" & Trim(!lotNumber)
            
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbLot.ListCount > 0 Then
      cmbLot.ListIndex = 0
      cmbLot = cmbLot.List(0)
      'lblNumber = sLots(0, 0)
      lblNumber = sLots(0)
      bGoodLot = GetThisLot()
   Else
      ManageBoxes 0, 1
      cmbLot = "No Lots Have Been Recorded"
      lblRem = "0.000"
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




'Leave Public - Called from elsewhere

Public Function GetThisLot() As Byte
   On Error GoTo DiaErr1
   ManageBoxes 0
   cmdChg.Enabled = False
'   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF," _
'          & "LOTUNITCOST,LOTDATECOSTED,LOTADATE,LOTMATLENGTH," _
'          & "LOTMATLENGTHUM,LOTMATHEIGHT,LOTMATHEIGHTHUM,LOTREMAININGQTY," _
'          & "LOTMATWIDTH,LOTMATWIDTHHUM,LOTLOCATION,LOTCOMMENTS,LOTSPLITCOMMENT," _
'          & "LOTSPLITFROMSYS,LOINUMBER,LOIRECORD," _
'          & "LOITYPE, LOTEXPIRESON FROM LohdTable,LoitTable WHERE " _
'          & "(LOTNUMBER='" & Trim(lblNumber) & " ' AND LOTNUMBER=LOINUMBER " _
'          & "AND LOIRECORD=1)"
   
   sSql = "SELECT LohdTable.*," & vbCrLf _
      & "LOINUMBER, LOIRECORD, LOITYPE" & vbCrLf _
      & "FROM LohdTable" & vbCrLf _
      & "JOIN LoitTable ON LOTNUMBER=LOINUMBER" & vbCrLf _
      & "WHERE LOTNUMBER='" & Trim(lblNumber) & "' AND LOIRECORD=1" & vbCrLf _
      & "AND LOTPARTREF = '" & Compress(Me.cmbPrt) & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_KEYSET)
   If bSqlRows Then
      ManageBoxes 1
      With RdoCur
         lblNumber = "" & Trim(!lotNumber)
         cmbLot = "" & Trim(!LOTUSERLOTID)
         txtCmt = "" & Trim(!LOTCOMMENTS)
         txtLoc = "" & Trim(!LOTLOCATION)
         txtCert = "" & !LOTCERT
         If IsNull(!LotUnitCost) Then
            txtUnitCost = Format(0, ES_QuantityDataFormat)
            cOldCost = 0
         Else
            txtUnitCost = Format(!LotUnitCost, ES_QuantityDataFormat)
            cOldCost = !LotUnitCost
         End If
         'If Not IsNull(.rdoColumns(6)) Then
         If Not IsNull(!LotADate) Then
            lblDate = "" & Format(!LotADate, "mm/dd/yyyy")
         Else
            lblDate = Format(GetServerDateTime, "mm/dd/yyyy")
         End If
         txtHgt = Format(!LOTMATHEIGHT, ES_QuantityDataFormat)
         txtLng = Format(!LOTMATLENGTH, ES_QuantityDataFormat)
         txtWid = Format(!LOTMATWIDTH, ES_QuantityDataFormat)
         If Val(txtHgt) > 0 Then txtHum = "" & Trim(!LOTMATHEIGHTHUM)
         If Val(txtLng) > 0 Then txtLum = "" & Trim(!LOTMATLENGTHUM)
         If Val(txtWid) > 0 Then txtWum = "" & Trim(!LOTMATWIDTHHUM)
         lblRem = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
         lblType = GetLotType(!LOITYPE)
         
         Me.lblQty = !LOTORIGINALQTY
         Me.txtTotMatl = !LOTTOTMATL
         Me.txtTotLabor = !LOTTOTLABOR
         Me.txtTotExp = !LOTTOTEXP
         Me.txtTotOh = !LOTTOTOH
         CalculateCostTotals
         
         If IsNull(!LOTEXPIRESON) Then
            cboExpirationDate.Text = ""
         Else
            cboExpirationDate = Format(!LOTEXPIRESON, "mm/dd/yyyy")
         End If
         
         If LotsExpire Then
            If cboExpirationDate.Visible And !LOTREMAININGQTY > 0 Then
               cboExpirationDate.Enabled = True
            Else
               cboExpirationDate.Enabled = False
            End If
         End If
         If Trim(!LOTSPLITFROMSYS) <> "" Then
            txtSplt = "" & Trim(!LOTSPLITCOMMENT)
            txtSplt.BackColor = cmbLot.BackColor
            txtSplt.Enabled = True
         Else
            txtSplt = ""
            txtSplt.BackColor = Me.BackColor
            txtSplt.Enabled = False
         End If
         sOldLot = lblNumber
      End With
      GetThisLot = 1
   Else
      ManageBoxes 0, 1
      GetThisLot = 0
      MsgBox "The request lot was not found or is not available.", _
         vbInformation, Caption
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthislot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ManageBoxes(bOpen As Byte, Optional BlankNumber As Byte)
   'Temp
   On Error Resume Next
   z1(16).Enabled = False
   z1(17).Enabled = False
   lblSon.Enabled = False
   cmbSon.Enabled = False
   cmbItm.Enabled = False
   
   z1(18).Enabled = False
   z1(19).Enabled = False
   cmbMon.Enabled = False
   cmbRun.Enabled = False
   
   'lblPart = ""
   'lblPart.ToolTipText = ""
   If BlankNumber = 1 Then lblNumber = ""
   lblType = ""
   lblRem = ""
   txtCmt = ""
   lblDate = ""
   txtHgt = "0.000"
   txtLng = "0.000"
   txtWid = "0.000"
   txtHum = ""
   txtLum = ""
   txtWum = ""
   txtLoc = ""
   lblDate = ""
   txtSplt = ""
   On Error Resume Next
   
   'Open the bottom for use
   If bOpen = 1 Then
      cmbLot.Enabled = True
      optMal.Enabled = True
      optSal.Enabled = True
      txtUnitCost.Enabled = True
      txtCmt.Enabled = True
      lblDate.Enabled = True
      txtHgt.Enabled = True
      txtLng.Enabled = True
      txtWid.Enabled = True
      txtHum.Enabled = True
      txtLum.Enabled = True
      txtWum.Enabled = True
      txtLoc.Enabled = True
      
   Else
      'open the top for use
      cmdChg.Enabled = False
      optMal.Enabled = False
      optSal.Enabled = False
      txtUnitCost.Enabled = False
      txtCmt.Enabled = False
      lblDate.Enabled = False
      txtHgt.Enabled = False
      txtLng.Enabled = False
      txtWid.Enabled = False
      txtHum.Enabled = False
      txtLum.Enabled = False
      txtWum.Enabled = False
      txtLoc.Enabled = False
      txtSplt.Enabled = False
      txtSplt.BackColor = Me.BackColor
   End If
   
End Sub

Private Sub txtTotMatl_LostFocus()
   If IsNumeric("0" & txtTotMatl) Then
      txtTotMatl.ForeColor = vbBlack
      With RdoCur
         If !LOTTOTMATL <> CCur("0" & txtTotMatl) Then
               !LOTTOTMATL = CCur("0" & txtTotMatl)
               txtTotMatl = !LOTTOTMATL
               .Update
            CalculateCostTotals
         End If
      End With
   Else
      txtTotMatl.ForeColor = vbRed
      txtTotMatl.SetFocus
   End If
End Sub

Private Sub txtTotLabor_LostFocus()
   If IsNumeric("0" & txtTotLabor) Then
      txtTotLabor.ForeColor = vbBlack
      With RdoCur
         If !LOTTOTLABOR <> CCur("0" & txtTotLabor) Then
            !LOTTOTLABOR = CCur("0" & txtTotLabor)
            txtTotLabor = !LOTTOTLABOR
            .Update
         CalculateCostTotals
         End If
      End With
   Else
      txtTotLabor.ForeColor = vbRed
      txtTotLabor.SetFocus
   End If
End Sub

Private Sub txtTotExp_LostFocus()
   If IsNumeric("0" & txtTotExp) Then
      txtTotExp.ForeColor = vbBlack
      With RdoCur
         If !LOTTOTEXP <> CCur("0" & txtTotExp) Then
            !LOTTOTEXP = CCur("0" & txtTotExp)
            txtTotExp = !LOTTOTEXP
            .Update
         CalculateCostTotals
         End If
      End With
   Else
      txtTotExp.ForeColor = vbRed
      txtTotExp.SetFocus
   End If
End Sub

Private Sub txtTotOh_LostFocus()
   If IsNumeric("0" & txtTotOh) Then
      txtTotOh.ForeColor = vbBlack
      With RdoCur
         If !LOTTOTOH <> CCur("0" & txtTotOh) Then
            !LOTTOTOH = CCur("0" & txtTotOh)
            txtTotOh = !LOTTOTOH
            .Update
         CalculateCostTotals
         End If
      End With
   Else
      txtTotOh.ForeColor = vbRed
      txtTotOh.SetFocus
   End If
End Sub

Private Sub txtUnitCost_LostFocus()
'   txtUnitCost = CheckLen(txtUnitCost, 9)
'   txtUnitCost = Format(Abs(Val(txtUnitCost)), ES_QuantityDataFormat)
'   If bGoodLot = 1 Then
'      'On Error Resume Next
'      If Val(txtUnitCost) <> cOldCost Then
'         With RdoCur
'            .Edit
'            !LOTUNITCOST = Format(Val(txtUnitCost), ES_QuantityDataFormat)
'            If Val(txtUnitCost) > 0 Then
'               !LOTDATECOSTED = Format(ES_SYSDATE, "mm/dd/yy")
'            Else
'               !LOTDATECOSTED = Null
'            End If
'            .Update
'         End With
'         cOldCost = Val(txtUnitCost)
'      End If
'   End If
'   CalculateCostTotals
   
   If IsNumeric("0" & txtUnitCost) Then
      txtUnitCost.ForeColor = vbBlack
      With RdoCur
         If !LotUnitCost <> CCur("0" & txtUnitCost) Then
            !LotUnitCost = CCur("0" & txtUnitCost)
            cOldCost = !LotUnitCost
            txtUnitCost = !LotUnitCost
            If CCur(txtUnitCost) > 0 Then
               !LOTDATECOSTED = Format(ES_SYSDATE, "mm/dd/yyyy")
            Else
               !LOTDATECOSTED = Null
            End If
            .Update
         CalculateCostTotals
         End If
      End With
   Else
      txtUnitCost.ForeColor = vbRed
      txtUnitCost.SetFocus
   End If
   
End Sub


Private Sub txtHgt_LostFocus()
   txtHgt = CheckLen(txtHgt, 8)
   txtHgt = Format(Abs(Val(txtHgt)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATHEIGHT = txtHgt
         .Update
      End With
   End If
   
End Sub


Private Sub txtHum_LostFocus()
   txtHum = CheckLen(txtHum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATHEIGHTHUM = txtHum
         .Update
      End With
   End If
   
End Sub


Private Sub txtLng_LostFocus()
   txtLng = CheckLen(txtLng, 8)
   txtLng = Format(Abs(Val(txtLng)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATLENGTH = txtLng
         .Update
      End With
   End If
   
End Sub


Private Sub txtCert_LostFocus()
   txtCert = CheckLen(txtCert, 40)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTCERT = txtCert
         .Update
      End With
   End If
   
End Sub

Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 4)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTLOCATION = txtLoc
         .Update
      End With
   End If
   
End Sub


Private Sub txtLum_LostFocus()
   txtLum = CheckLen(txtLum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATLENGTHUM = txtLum
         .Update
      End With
   End If
   
End Sub


Private Sub txtSplt_LostFocus()
   txtSplt = CheckLen(txtSplt, 20)
   txtSplt = StrCase(txtSplt, ES_FIRSTWORD)
   On Error Resume Next
   With RdoCur
      !LOTSPLITCOMMENT = txtSplt
      .Update
   End With
   
End Sub


Private Sub txtWid_LostFocus()
   txtWid = CheckLen(txtWid, 8)
   txtWid = Format(Abs(Val(txtWid)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATWIDTH = txtWid
         .Update
      End With
   End If
   
End Sub


Private Sub txtWum_LostFocus()
   txtWum = CheckLen(txtWum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATWIDTHHUM = txtWum
         .Update
      End With
   End If
   
End Sub



Private Function GetLotPart(sLotPart As String) As Byte
   Dim RdoPrt As ADODB.Recordset
   cmbLot.Clear
   'sSql = "Qry_GetPartsNotTools '" & Compress(cmbPrt) & "'"
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PATOOL, PALOTSEXPIRE FROM PartTable" & vbCrLf _
      & "WHERE PARTREF = '" & Compress(cmbPrt) & "' AND PATOOL=0"
   If clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD) Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         LotsExpire = IIf(!PALOTSEXPIRE = 0, False, True)
         GetLotPart = 1
         ClearResultSet RdoPrt
         'LotsExpire = IIf(!PALOTSEXPIRE = 0, False, True)
      End With
   Else
      LotsExpire = False
      lblDsc = "Part Number With Lot Wasn't Found."
      GetLotPart = 0
   End If
   lblExpirationDate.Visible = LotsExpire
   cboExpirationDate.Visible = LotsExpire
   
   If GetLotPart = 1 Then GetLots
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlotpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetLotType(bType As Byte) As String
   Select Case bType
      Case 15
         GetLotType = "Purchase Order Receipt"
      Case 6
         GetLotType = "MO Completion"
      Case 19
         GetLotType = "Manual Adjustment"
      Case Else
         GetLotType = "Other Inventory Adustment"
   End Select
   
End Function

Private Sub cboExpirationDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub CalculateCostTotals()
   
   'lblExtCost = CCur("0" & lblQty) * CCur("0" & txtUnitCost)
   lblExtCost = String2Currency(lblQty) * String2Currency(txtUnitCost)
   lblTotCost = String2Currency(txtTotMatl) + String2Currency(txtTotLabor) _
      + String2Currency(txtTotExp) + String2Currency(txtTotOh)
      
   If lblExtCost.Caption = lblTotCost.Caption Then
      lblExtCost.ForeColor = vbBlack
      lblTotCost.ForeColor = vbBlack
   Else
      lblExtCost.ForeColor = vbRed
      lblTotCost.ForeColor = vbRed
   End If
End Sub

Private Function CostsAreBalanced() As Boolean
   If lblExtCost.Caption <> lblTotCost.Caption Then
      MsgBox "The extended and total costs do not agree.  Please change the costs?", vbCritical
      CostsAreBalanced = False
'      Case vbYes
'      Case Else
'         'txtUnitCost.SetFocus
'         Exit Function
'      End Select
   Else
      CostsAreBalanced = True
   End If
End Function
