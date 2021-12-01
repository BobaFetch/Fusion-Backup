VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Invoice"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "diaARe01b.frx":0000
      DownPicture     =   "diaARe01b.frx":0972
      Height          =   350
      Left            =   6960
      Picture         =   "diaARe01b.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Standard Comments"
      Top             =   4200
      Width           =   350
   End
   Begin VB.TextBox txtNet 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CheckBox optCan 
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkEOM 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Age From End Of Month"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtDays 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Tag             =   "1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtARDisc 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSav 
      Caption         =   "Saved"
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optItm 
      Caption         =   "Items"
      Height          =   195
      Left            =   5520
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Height          =   315
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   "List Sales Order Items"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtFrt 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox optLoad 
      Caption         =   "Load"
      Height          =   195
      Left            =   5520
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1695
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   "9"
      Top             =   4200
      Width           =   5295
   End
   Begin VB.TextBox txtStAdr 
      Height          =   1155
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   9
      Tag             =   "9"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtCar 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1560
      Width           =   1875
   End
   Begin VB.TextBox txtWay 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1560
      Width           =   1875
   End
   Begin VB.ComboBox txtShd 
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
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
      PictureUp       =   "diaARe01b.frx":1C56
      PictureDn       =   "diaARe01b.frx":1D9C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6030
      FormDesignWidth =   7455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "EOM"
      Height          =   255
      Index           =   14
      Left            =   4200
      TabIndex        =   34
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   33
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   12
      Left            =   2280
      TabIndex        =   32
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Terms"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Comments:"
      Height          =   495
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship To:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carrier"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   24
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Way Bill"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Date"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   22
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   20
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   19
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "diaARe01b"
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
'
' diaARe01ah - Revise sales order invoice header
'
' Created: (cjs)
' Revisions:
'   06/13/02 (nth) Added INVARDISC and INVDAYS.
'   06/14/02 (nth) Added EOM
'   08/20/02 (nth) Correctly format INVSHIPDATE and INVDATE m/d/yyyy
'   08/26/02 (nth) Added INVNETDAYS per MCS
'   10/28/03 (nth) check for open journal by invoice date
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Public bGoodInvoice As Boolean
Dim sMsg As String

Public sJournalID As String
Public RdoInv As ADODB.Recordset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub chkEOM_Click()
   If bGoodInvoice Then
      On Error Resume Next
      If chkEOM Then
         RdoInv!INVEOM = 1
      Else
         RdoInv!INVEOM = 0
      End If
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdComments_Click()
'bbs changes from the Comments form to the SysComments form on 6/28/2010 for Ticket #31511
   'Add one of these to the form and go
   If cmdComments Then
      'The Default is txtCmt and need not be included
      'Use Select Case cmdCopy to add your own
      txtCmt.SetFocus
      SysComments.lblControl = "txtCmt"
      'See List For Index
      SysComments.lblListIndex = 4
      SysComments.Show
      cmdComments = False
   End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Customer Invoice (Sales Order)"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdItm_Click()
   sJournalID = GetOpenJournal("SJ", Format(txtDte, "mm/dd/yy"))
   If sJournalID = "" Then
      sMsg = "There Is No Open Journal For The Posting Date."
      MsgBox sMsg, vbInformation, Caption
      txtDte.SetFocus
      Exit Sub
   End If
   optItm = vbChecked
   diaARe01c.lblSon = lblSon
   diaARe01c.lblCst = lblCst
   diaARe01c.lblNme = lblNme
   diaARe01c.Show
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'bGoodInvoice = GetInvoice
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me ' ES_DONTLIST, ES_RESIZE
   
   'Move diaARe01a.Top + 200, diaARe01a.Left + 200
   FormatControls
   sCurrForm = Caption
   txtDte = Format(Now, "mm/dd/yyyy")
   txtShd = Format(Now, "mm/dd/yyyy")
   txtFrt = "0.00"
   bOnLoad = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   optCan.Value = vbChecked
   
   If optItm.Value = vbChecked Then
      Unload diaARe01c
   End If
   
   ' Invoice has no items, Cancel it?
   
   If optSav.Value = vbUnchecked Then
      MsgBox "There Were No Items Recorded. " & vbCrLf _
         & "The Invoice Will be Canceled.", vbInformation, Caption
      
      clsADOCon.BeginTrans
      sSql = "DELETE FROM CihdTable WHERE INVNO=" & lCurrInvoice & " "
      clsADOCon.ExecuteSQL sSql
      Dim inv As New ClassARInvoice
      inv.SaveLastInvoiceNumber lCurrInvoice - 1
      clsADOCon.CommitTrans
   End If
   lCurrInvoice = 0
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoInv = Nothing
   diaARe01a.Show
   Set diaARe01b = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optCan_Click()
   ' never visible, used to tell when form is unloading
End Sub

Private Sub optItm_Click()
   ' Never visible - Items
End Sub

Private Sub optLoad_Click()
   ' Never visible - Files:
   GetCustomerInfo
   bGoodInvoice = GetInvoice
End Sub


Private Sub optSav_Click()
   'Never visible
   'If not checked then the invoice will be dumped
End Sub

Private Sub txtARDisc_LostFocus()
   txtARDisc = CheckLen(txtARDisc, 8)
   txtARDisc = Format(Abs(Val(txtARDisc)), "####0.00")
   If bGoodInvoice Then
      RdoInv!INVARDISC = Val(txtARDisc)
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtCar_LostFocus()
   txtCar = CheckLen(txtCar, 20)
   If bGoodInvoice Then
      RdoInv!INVCARRIER = txtCar
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = CheckComments(txtCmt)
   If bGoodInvoice Then
      RdoInv!INVCOMMENTS = Trim(txtCmt)
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtDays_LostFocus()
   If bGoodInvoice Then
      RdoInv!INVDAYS = Val(txtDays)
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   txtShd = txtDte
   CurrentJournal "SJ", txtDte, sJournalID
   If bGoodInvoice Then
      RdoInv!INVDATE = Format(txtDte, "mm/dd/yyyy")
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 8)
   txtFrt = Format(Abs(Val(txtFrt)), "####0.00")
   If bGoodInvoice Then
      RdoInv!INVFREIGHT = Val(txtFrt)
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtNet_LostFocus()
   If bGoodInvoice Then
      On Error Resume Next
      RdoInv!INVNETDAYS = Val(txtNet)
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtShd_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtShd_LostFocus()
   txtShd = CheckDateEx(txtShd)
   If txtShd < txtDte Then
      Beep
      txtShd = Format(Now, "mm/dd/yyyy")
   End If
   If bGoodInvoice Then
      On Error Resume Next
      RdoInv!INVSHIPDATE = Format(txtShd, "mm/dd/yyyy")
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtStAdr_LostFocus()
   txtStAdr = CheckLen(txtStAdr, 255)
   If bGoodInvoice Then
      RdoInv!INVSTADR = txtStAdr
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub


Private Sub txtWay_LostFocus()
   txtWay = CheckLen(txtWay, 16)
   If bGoodInvoice Then
      RdoInv!INVWAYBILL = txtWay
      RdoInv.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Public Sub GetCustomerInfo()
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOVIA,SODAYS,SOARDISC,SONETDAYS FROM SohdTable " _
          & "WHERE SONUMBER=" & Val(Right(lblSon, SO_NUM_SIZE))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      On Error Resume Next
      With RdoSon
         txtCar = "" & Trim(!SOVIA)
         txtDays = "" & Trim(!SODAYS)
         txtNet = "" & Trim(!SONETDAYS)
         txtARDisc = "" & Trim(!SOARDISC)
         .Cancel
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomerin"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetInvoice() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT INVNO,INVSHIPDATE,INVSTADR,INVCOMMENTS,INVNETDAYS," _
          & "INVTAX,INVFREIGHT,INVDATE,INVCARRIER,INVWAYBILL,INVARDISC," _
          & "INVDAYS,INVEOM FROM CihdTable WHERE INVNO=" & lCurrInvoice & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   
   If bSqlRows Then
      On Error Resume Next
      With RdoInv
         RdoInv!INVCARRIER = "" & Trim(txtCar)
         RdoInv!INVDAYS = Val(txtDays)
         RdoInv!INVNETDAYS = Val(txtNet)
         RdoInv!INVARDISC = Val(txtARDisc)
         RdoInv!INVDATE = Format(txtDte, "mm/dd/yyyy")
         RdoInv!INVSHIPDATE = Format(txtShd, "mm/dd/yyyy")
         RdoInv.Update
         If Err > 0 Then
            ValidateEdit Me
            GetInvoice = 0
         Else
            GetInvoice = 1
         End If
      End With
   Else
      GetInvoice = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
