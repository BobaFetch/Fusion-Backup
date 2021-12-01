VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map ES/2000 ERP Customers To QuickBooks ®"
   ClientHeight    =   1875
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQBName 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   3000
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1560
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1875
      FormDesignWidth =   5400
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
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
      PictureUp       =   "diaARe09a.frx":0000
      PictureDn       =   "diaARe09a.frx":0146
   End
   Begin VB.Label lblnme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ES/2002 Customer"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   375
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickBooks ® Customer"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1065
   End
End
Attribute VB_Name = "diaARe09a"
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

'*************************************************************************************
' diaARe09a - Map ES/2002 Customers To QuickBooks
'
' Created: 06/18/02 (nth)
' Revisions:
'   11/05/02 (nth) Increased checklen of qb name from 30 to 50
'
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bGoodCust As Byte
Dim rdoCst As ADODB.Recordset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   bGoodCust = GetCustomer
End Sub

Private Sub cmbCst_LostFocus()
   FindCustomer Me, cmbCst
   bGoodCust = GetCustomer
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = True
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentCustomer = cmbCst
   FormUnload
   Set rdoCst = Nothing
   Set diaARe09a = Nothing
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillCustomers Me
   If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
   FindCustomer Me, cmbCst
   bGoodCust = GetCustomer()
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Function GetCustomer() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT CUQBNAME FROM CustTable " _
          & "WHERE CUREF = '" & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_KEYSET)
   If bSqlRows Then
      txtQBName = "" & Trim(rdoCst!CUQBNAME)
      GetCustomer = 1
   Else
      txtQBName = ""
      GetCustomer = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtQBName_LostFocus()
   txtQBName = CheckLen(txtQBName, 50)
   If bGoodCust Then
      On Error Resume Next
      rdoCst!CUQBNAME = "" & Trim(txtQBName)
      rdoCst.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub
