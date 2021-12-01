VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Parts To Buyers"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Height          =   315
      Left            =   5565
      TabIndex        =   15
      ToolTipText     =   "Next Page"
      Top             =   2880
      Width           =   875
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      ToolTipText     =   "Last Page"
      Top             =   2880
      Width           =   875
   End
   Begin VB.CommandButton cmdAsn 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Assign The Selected Buyer To The Current Part Number"
      Top             =   2160
      Width           =   875
   End
   Begin VB.ComboBox cmbByr 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select A Buyer "
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Select A Product Code"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Select Part Numbers"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Leading Char(s) Fills Up To 300 Part Numbers"
      Top             =   1200
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3390
      FormDesignWidth =   6510
   End
   Begin VB.Label lblParts 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5520
      TabIndex        =   13
      ToolTipText     =   "Total Row Count"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label lblCurBuyer 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Buyer"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer ID"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   3315
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1305
   End
End
Attribute VB_Name = "PurcPRe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim iIndex As Integer

Dim sPartNumbers(301, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   cmbPrt.Clear
   lblParts = "0"
   
End Sub

Private Sub cmbCde_LostFocus()
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbPrt_Click()
   cmdAsn.Enabled = True
   If iIndex > -1 Then
      iIndex = cmbPrt.ListIndex
      lblDsc = sPartNumbers(iIndex, 1)
      lblCurBuyer = sPartNumbers(iIndex, 2)
   End If
   
End Sub


Private Sub cmdAsn_Click()
   cmdAsn.Enabled = False
   On Error Resume Next
   sSql = "UPDATE PartTable SET PABUYER='" & Compress(cmbByr) & "' " _
          & "WHERE PARTREF='" & Compress(cmbPrt) & "'"
   clsADOCon.ExecuteSQL sSql
   lblCurBuyer = cmbByr
   sPartNumbers(iIndex, 2) = cmbByr
   
   ' Update the Association Table BuyerPart table
   UpdateBuyerPart Compress(cmbPrt), Compress(cmbByr)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4307
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLst_Click()
   cmdAsn.Enabled = True
   iIndex = iIndex - 1
   If iIndex < 0 Then iIndex = 0
   cmbPrt.ListIndex = iIndex
   cmbPrt_Click
   
End Sub

Private Sub cmdNxt_Click()
   On Error GoTo DiaErr1
   cmdAsn.Enabled = True
   iIndex = iIndex + 1
   If iIndex > cmbPrt.ListCount - 1 Then iIndex = cmbPrt.ListCount - 1
   cmbPrt.ListIndex = iIndex
   cmbPrt_Click
   Exit Sub
DiaErr1:
   On Error GoTo 0
   
End Sub


Private Sub cmdSel_Click()
   FillParts
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillBuyers
      If cmbByr.ListCount > 0 Then
         cmbByr = cmbByr.List(0)
         GetCurrentBuyer cmbByr, 1
      Else
         lblCurBuyer = "*** No Buyers Found ***"
      End If
      cmbCde.AddItem "ALL"
      FillProductCodes
      FillParts
      bOnLoad = 0
   End If
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
   Set PurcPRe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub FillParts()
   Dim RdoPrt As ADODB.Recordset
   Dim sCode As String
   iIndex = -1
   Erase sPartNumbers
   If cmbCde <> "ALL" Then sCode = Compress(cmbCde)
   On Error GoTo DiaErr1
'   sSql = "select PARTREF,PARTNUM,PADESC,PAPRODCODE,PABUYER,BYREF,BYNUMBER " _
'          & "FROM PartTable,BuyrTable WHERE (PABUYER*=BYREF AND " _
'          & "PARTREF LIKE '" & Compress(cmbPrt) & "%' AND PAPRODCODE " _
'          & "LIKE '" & sCode & "%')"
   sSql = "select PARTREF,PARTNUM,PADESC,PAPRODCODE,PABUYER,BYREF,BYNUMBER " _
      & "FROM PartTable" & vbCrLf _
      & "LEFT JOIN BuyrTable ON PABUYER=BYREF" & vbCrLf _
      & "WHERE PARTREF LIKE '" & Compress(cmbPrt) & "%' AND PAPRODCODE " _
      & "LIKE '" & sCode & "%' AND PAINACTIVE = 0 AND PAOBSOLETE = 0"
   cmbPrt.Clear
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         Do Until .EOF
            iIndex = iIndex + 1
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            sPartNumbers(iIndex, 0) = "" & Trim(!PartNum)
            sPartNumbers(iIndex, 1) = "" & Trim(!PADESC)
            sPartNumbers(iIndex, 2) = "" & Trim(!BYNUMBER)
            If iIndex > 299 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoPrt
      End With
   End If
   lblParts = iIndex + 1
   If cmbPrt.ListCount > 0 Then
      iIndex = 0
      If cmbByr.ListCount > 0 Then cmdAsn.Enabled = True
      cmbPrt.ListIndex = iIndex
      cmbPrt = sPartNumbers(iIndex, 0)
      lblDsc = sPartNumbers(iIndex, 1)
      lblCurBuyer = sPartNumbers(iIndex, 2)
   Else
      cmdAsn.Enabled = False
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblCurBuyer_Change()
   If Left(lblCurBuyer, 6) = "*** No" Then _
           lblCurBuyer.ForeColor = ES_RED Else _
           lblCurBuyer.ForeColor = vbBlack
   
End Sub


Private Sub UpdateBuyerPart(sPartNum As String, sBuyer As String)
    
    On Error GoTo DiaErr1
    If (Len(sPartNum) > 0 And Len(sBuyer) > 0) Then
        
         ' Delete the existing part/buyer association
         sSql = "DELETE FROM BuypTable WHERE " _
                & "BYPARTNUMBER='" & sPartNum & "'"
         
         clsADOCon.ExecuteSQL sSql
         
         
         ' Add the new part/buyer association
         sSql = "INSERT INTO BuypTable (BYREF,BYPARTNUMBER) " _
                & "VALUES('" & sBuyer & "','" & sPartNum & "')"
         clsADOCon.ExecuteSQL sSql
        
        Exit Sub
        
DiaErr1:
        sProcName = "GetProductCode"
        CurrError.Number = Err.Number
        CurrError.Description = Err.Description
        DoModuleErrors Me
        
    End If
End Sub



