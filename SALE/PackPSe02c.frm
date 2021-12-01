VERSION 5.00
Begin VB.Form PackPSe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Sales Order Items To PackSlip"
   ClientHeight    =   5430
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   2340
      Width           =   3015
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe02c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      Top             =   2340
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox txtBook 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Tag             =   "4"
      ToolTipText     =   "This Item Was Booked"
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PackPSe02c.frx":07AE
      Height          =   315
      Left            =   4560
      Picture         =   "PackPSe02c.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   2340
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox txtDdt 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Tag             =   "4"
      ToolTipText     =   "Expected Actual Delivery (With Freight Days)"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.ComboBox txtRdt 
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Tag             =   "4"
      ToolTipText     =   "Customer Requested Date"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.ComboBox txtSdt 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Tag             =   "4"
      ToolTipText     =   "Expected Ship Date"
      Top             =   3780
      Width           =   1095
   End
   Begin VB.TextBox txtDiscountPercent 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Discount Percentage"
      Top             =   3780
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   5760
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Add This Sales Order Item"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtListPrice 
      Height          =   285
      Left            =   300
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Unit Price"
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   34
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   33
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblSon 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   32
      ToolTipText     =   "Our Item Number"
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   31
      ToolTipText     =   "Our Item Number"
      Top             =   480
      Width           =   315
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   29
      ToolTipText     =   "Our Item Number"
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblDiscountAmount 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   300
      TabIndex        =   28
      ToolTipText     =   "Quantity On Hand"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label lblNetPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2340
      TabIndex        =   27
      ToolTipText     =   "Quantity On Hand"
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
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
      Index           =   17
      Left            =   1440
      TabIndex        =   26
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price                 "
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
      Index           =   16
      Left            =   2400
      TabIndex        =   25
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   23
      Top             =   4920
      Width           =   1050
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   300
      TabIndex        =   22
      ToolTipText     =   "Quantity On Hand"
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   21
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   20
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Date        "
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
      Index           =   5
      Left            =   3480
      TabIndex        =   19
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      ToolTipText     =   "Our Item Number"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity/Qoh            "
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
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                           "
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
      Index           =   3
      Left            =   1500
      TabIndex        =   15
      Top             =   2100
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Price/Disc"
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
      Index           =   4
      Left            =   300
      TabIndex        =   14
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1440
      TabIndex        =   13
      ToolTipText     =   "Extended Description"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1500
      TabIndex        =   12
      ToolTipText     =   "Item Revision"
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "PackPSe02c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit

Dim strOldPart As String
Dim bOnLoad As Byte


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdAdd_Click()
   If (txtPrt <> "") Then
      'MdiSect.ActiveForm.txtTest = "Test1"
      If (AddSoItem = True) Then
         Dim iNewItem As Integer
         iNewItem = Val(lblItm)
         PackPSe02b.AddNewSOtoPS iNewItem
      Else
         MsgBox "Could not add new Sales Item.", vbInformation, Caption
      End If
      ' Unload the form
      Unload Me
      
   Else
         MsgBox "Please select Part Number.", vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = txtPrt
   'optVew.Value = vbChecked
   ViewParts.Show
End Sub


Private Function GetPart(strGetPart As String) As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sComment As String
   
   On Error GoTo DiaErr1
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            'cmbPrt = "" & Trim(!PARTNUM)
            'txtPrt = "" & Trim(!PARTNUM)
            lblDsc = "" & Trim(!PADESC)
            If strOldPart <> txtPrt Then txtListPrice = Format(!PAPRICE, ES_QuantityDataFormat)
            lblQoh = Format(!PAQOH, ES_QuantityDataFormat)
            strOldPart = txtPrt
            GetPart = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPart = 0
         cmbPrt = ""
         ' Don't reset txt for partial searching.
         txtPrt = ""
         lblDsc = "*** Part Wasn't Found ***"
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      txtPrt = ""
      cmbPrt = ""
      lblDsc = ""
   End If
   lblDsc = lblDsc & sComment
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmdTrm_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim lItem As Long
   Dim bPartSearch As Boolean
   lItem = GetLastItem
   lblItm = lItem
   lblRev = GetSORev
   
   bPartSearch = GetPartSearchOption
   SetPartSearchOption (bPartSearch)
   If (Not bPartSearch) Then FillPartCombo cmbPrt
   cmbPrt = ""
   ' Add todays date
   txtSdt = Format(GetServerDateTime, "mm/dd/yy")
   txtBook = Format(GetServerDateTime, "mm/dd/yy")
   txtDdt = Format(GetServerDateTime, "mm/dd/yy")
   txtRdt = Format(GetServerDateTime, "mm/dd/yy")
   bOnLoad = 0
End Sub

Private Sub Form_Load()

   ES_SellingPriceFormat = GetSellingPriceFormat()
   bOnLoad = 1
End Sub


'10/27/04

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'On Error Resume Next
   Set PackPSe02c = Nothing
   
End Sub

Private Function GetLastItem() As Long
   Dim RdoLst As ADODB.Recordset
   'On Error Resume Next
   sSql = "SELECT MAX(ITNUMBER) FROM SoitTable WHERE " _
          & "ITSO=" & Val(lblSon) & ""
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         If Not IsNull(.Fields(0)) Then
            GetLastItem = .Fields(0) + 1
         Else
            GetLastItem = 1
         End If
         ClearResultSet RdoLst
      End With
   Else
      GetLastItem = 1
   End If
   Set RdoLst = Nothing
   If Err > 0 Then GetLastItem = 1
   
End Function

Private Function GetSORev() As String
   Dim RdoLst As ADODB.Recordset
   'On Error Resume Next
   sSql = "SELECT ITREV FROM SoitTable WHERE " _
          & "ITSO=" & Val(lblSon) & ""
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         If Not IsNull(.Fields(0)) Then
            GetSORev = .Fields(0)
         Else
            GetSORev = ""
         End If
         ClearResultSet RdoLst
      End With
   Else
      GetSORev = ""
   End If
   Set RdoLst = Nothing
   If Err > 0 Then GetSORev = 1
   
End Function


Private Sub txtDiscountPercent_LostFocus()
   CalculateDiscount
End Sub

Private Sub txtListPrice_LostFocus()
   txtListPrice = Format(Val(txtListPrice), ES_SellingPriceFormat)
   CalculateDiscount
End Sub

'Private Sub txtPrt_Change()
'   cmbPrt = txtPrt
'End Sub

Private Sub cmbPrt_Click()
    txtPrt = cmbPrt

End Sub

Private Sub cmbPrt_Change()
    txtPrt = cmbPrt
End Sub

Private Sub txtPrt_LostFocus()
   'On Error Resume Next
'   GetPart (txtPrt)
   'cmbPrt = txtPrt
End Sub

Private Sub cmbPrt_LostFocus()
    txtPrt = cmbPrt
   GetPart (cmbPrt)
   
End Sub


Private Sub txtSdt_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtSdt_LostFocus()
   txtSdt = CheckDate(txtSdt)
End Sub

Private Sub txtRdt_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtRdt_LostFocus()
   txtRdt = CheckDate(txtRdt)
End Sub

Private Sub txtDdt_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtDdt_LostFocus()
   txtDdt = CheckDate(txtDdt)
End Sub


Private Sub txtBook_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBook_LostFocus()
   txtBook = CheckDate(txtBook)
End Sub

Public Sub CalculateDiscount()
   Dim listPrice As Currency, discRate As Currency, discAmount As Currency
   'listPrice = "0" & Me.txtListPrice
   listPrice = IIf(IsNumeric(txtListPrice), txtListPrice, 0)
   'discRate = "0" & Me.txtDiscountPercent
   discRate = IIf(IsNumeric(txtDiscountPercent), txtDiscountPercent, 0)
   discAmount = listPrice * discRate / 100
   'Me.lblDiscountAmount = Format(discAmount, "######0.000")
   'Me.lblNetPrice = Format(CCur(listPrice) - discAmount, "######0.0000")
   Me.lblDiscountAmount = Format(discAmount, ES_SellingPriceFormat)
   Me.lblNetPrice = Format(CCur(listPrice) - discAmount, ES_SellingPriceFormat)
End Sub

Private Function GetSellingPriceFormat() As String
   Dim RdoSpf As ADODB.Recordset
   On Error GoTo 0
   sSql = "SELECT SellingPriceFormat FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSpf, ES_FORWARD)
   If bSqlRows Then GetSellingPriceFormat = RdoSpf!SellingPriceFormat _
                                            Else GetSellingPriceFormat = "#######0.000"
   Set RdoSpf = Nothing
   Exit Function
modErr1:
   GetSellingPriceFormat = "#######0.000"
   
End Function

Private Function AddSoItem()
   Dim iNewItem As Integer
   Dim iRows As Integer
   Dim iDelDays As Integer
   
   Dim lCurr As Long
   Dim lNew As Long
   Dim strNewPart As String
   
   iNewItem = Val(lblItm)
   strNewPart = Compress(txtPrt)
   If Val(txtQty) = 0 Then txtQty = Format(0, ES_QuantityDataFormat)
   If Val(txtDiscountPercent) = 0 Then txtDiscountPercent = Format(0, ES_SellingPriceFormat)
   
   On Error GoTo DiaErr1
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0

   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITPART,ITQTY,ITCUSTREQ,ITSCHED,ITBOOKDATE,ITSCHEDDEL, ITUSER," _
          & "ITDOLLORIG, ITDOLLARS, ITDISCRATE, ITDISCAMOUNT) " _
          & "VALUES(" & lblSon & "," & iNewItem & ",'" _
          & strNewPart & "'," & Val(txtQty) & ",'" & Format(txtRdt, "mm/dd/yy") & "','" _
          & Format(txtSdt, "mm/dd/yy") & "','" & Format(txtBook, "mm/dd/yy") & "','" & Format(txtDdt, "mm/dd/yy") & "','" _
          & sInitials & "'," & CStr(txtListPrice) & "," _
          & CStr(lblNetPrice) & "," & CStr(txtDiscountPercent) & "," _
          & CStr(lblDiscountAmount) & ")"
          
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   clsADOCon.CommitTrans
   
   If clsADOCon.RowsAffected = 0 Then
      MsgBox "Couldn't Add Item.", vbInformation, Caption
      AddSoItem = False
   Else
      AddSoItem = True
   End If
      
   Exit Function
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

