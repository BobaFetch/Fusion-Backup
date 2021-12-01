VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form diaARe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Invoice (Packing Slip)"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   5160
      TabIndex        =   31
      Text            =   "1"
      ToolTipText     =   "Number of Invoices to Print (After Post)"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CheckBox chkInvPS 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5160
      TabIndex        =   29
      ToolTipText     =   "Create An Advanced Payment Invoice "
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "diaARe02a.frx":0000
      DownPicture     =   "diaARe02a.frx":0972
      Height          =   350
      Left            =   5760
      Picture         =   "diaARe02a.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Standard Comments"
      Top             =   3600
      Width           =   350
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtFrt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Freight From Pack Slip (If Any) Or Enter Freight"
      Top             =   2880
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar Prg1 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtCmt 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "9"
      Top             =   4080
      Width           =   5895
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "diaARe02a.frx":1C56
      Height          =   315
      Left            =   3240
      Picture         =   "diaARe02a.frx":2130
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "View Selected Items"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Printed Pack Slips "
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbPre 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Invoice Prefix(A-Z)"
      Top             =   1560
      Width           =   510
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "P&ost"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   14
      ToolTipText     =   "Add This Packing Slip Invoice"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtInv 
      Height          =   315
      Left            =   2410
      TabIndex        =   2
      ToolTipText     =   "Requires A Number"
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   9
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
      PictureUp       =   "diaARe02a.frx":260A
      PictureDn       =   "diaARe02a.frx":2750
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6015
      FormDesignWidth =   6255
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   225
      Left            =   360
      TabIndex        =   26
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe02a.frx":2896
      PictureDn       =   "diaARe02a.frx":29DC
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies to Print:"
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice same as PS"
      Height          =   285
      Index           =   20
      Left            =   3720
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   27
      Top             =   0
      Width           =   2760
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   24
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblJrn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Comments:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Total"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Invoice Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   375
      Width           =   1695
   End
End
Attribute VB_Name = "diaARe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'***************************************************************************************
' diaARe02a - Customer Invoice From Packing Slip
'
' Created: (cjs)
'
' Revisions:
' 06/05/01 (nth) Make post date combo display calander on drop down.
' 06/07/01 (nth) Added support for COGS and INV accounts.
' 06/10/01 (nth) To update the ITREVACCT in sales order item table.
' 06/18/01 (nth) Cleaned up and tested.
' 06/25/01 (nth) Added tax and freight to invoice total.
' 07/31/01 (nth) Added option to print after packslip is invoiced.
' 11/05/01 (nth) Fixed error open journal was found window would not close.
' 12/19/02 (nth) Added B&O and sales tax logic.
' 01/08/02 (nth) Remeber last invoice prefix, requested by Jevco.
' 10/27/03 (nth) Added sales tax code logic improved BnO
' 12/04/03 (nth) Added item total and grand total
' 12/18/03 (nth) Clear tax varibles between packslip changes
' 02/27/04 (nth) Added LoitTable customer and invoice update
' 04/01/04 (nth) Removed the prompt to cancel the TM invoice.
'
'***************************************************************************************

Option Explicit

Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim AdoQry2 As ADODB.Command
Dim AdoParameter2 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPs As Boolean
Dim bInvUsed As Boolean
Dim bGoodAct As Boolean
Dim iTotalItems As Integer
Dim lNewInv As Long
Dim cFREIGHT As Currency
Dim cTax As Currency
Dim sAccount As String
Dim sInvPre As String
Dim sPsStadr As String
Dim sPsCust As String
Dim lSo As Long
Dim sMsg As String

' Sales Tax
Dim sTaxCode As String
Dim sTaxState As String
Dim sTaxAccount As String
Dim nTaxRate As Currency

' Sales journal
Dim sCOSjARAcct As String
Dim sCOSjINVAcct As String
Dim sCOSjNFRTAcct As String
Dim sCOSjTFRTAcct As String
Dim sCOSjTaxAcct As String

Dim vItems(300, 12) As Variant
'0 = SONUMBER
'1 = SOITEM
'2 = SOREV
'3 = Part Number (compressed)
'4 = Quantity
'5 = Account
'6 = total selling
'7 = Type (level)
'8 = Product Code
'9 = Standard Cost
'10 = Number
'11 = Selling Price
'12 = Tax Exempt

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'See if the Accounts are there

Public Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   If sJournalID = "" Then
      bGoodAct = True
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         ' A/R
         sCOSjARAcct = "" & Trim(.Fields(1))
         If sCOSjARAcct = "" Then b = 1
         ' NonTaxable freight
         sCOSjNFRTAcct = "" & Trim(.Fields(2))
         If sCOSjNFRTAcct = "" Then b = 1
         ' Taxable freight
         sCOSjTFRTAcct = "" & Trim(.Fields(3))
         If sCOSjTFRTAcct = "" Then b = 1
         ' Sales tax
         sCOSjTaxAcct = "" & Trim(.Fields(4))
         If sCOSjTaxAcct = "" Then b = 1
         .Cancel
      End With
   End If
   Set rdoJrn = Nothing
   If b = 1 Then
      bGoodAct = False
      lblJrn.Visible = True
   Else
      bGoodAct = True
      lblJrn.Visible = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbPre_LostFocus()
   cmbPre = CheckLen(cmbPre, 1)
   If Asc(cmbPre) < 65 Or Asc(cmbPre) > 90 Then
      Beep
      cmbPre = sInvPre
   End If
   sInvPre = cmbPre
End Sub

Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
End Sub

Private Sub cmbPsl_LostFocus()
   If Not bCancel Then
      cmbPsl = CheckLen(cmbPsl, 8)
      bGoodPs = GetPackslip()
      If bGoodPs Then
         GetPsItems
         ' Only if the same as PS
         If (chkInvPS.Value = 1) Then
            Dim strPreInv As String
            Dim strNewInv As String
            ' Store the old value
            strPreInv = txtInv
            GetNextInvoice
            strNewInv = txtInv
            If (Val(strNewInv) <> Val(strPreInv)) Then
               ' Delete the old temp Invoice created at load/change.
               DeleteOldTmpInv (strPreInv)
               ' Add new
               AddInvoice
            End If
            ' If the
         End If
      End If
   End If
End Sub

Private Sub cmdAdd_Click()
    Dim success As Boolean
    success = True
    
   ' check for open journal via posting date
   sJournalID = GetOpenJournal("SJ", Format(txtDte, "mm/dd/yy"))
   If sJournalID = "" Then
      sMsg = "There Is No Open Journal For The Posting Date."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   GetSJAccounts
   
   Dim bAllowReuse As Boolean
   
   If bGoodAct = False Then
      MsgBox "One Or More Journal Accounts Are Not Registered." & vbCr _
         & "Please Install All Accounts In the Company Setup.", _
         vbInformation, Caption
   Else
      If Val(txtInv) <> lNewInv Then
         bInvUsed = GetOldInvoice(txtInv)
         
         If (bInvUsed) Then
            'Check if the Invoice was cancelled.
            Dim strPackSlip As String
            
            bAllowReuse = CheckForCancelInv(txtInv, strPackSlip)
            If (strPackSlip <> "" And bAllowReuse = True) Then bAllowReuse = False
            
         End If
         
      Else
         bInvUsed = False
      End If
      
      If bInvUsed Then
         If (bAllowReuse) Then
            Dim bResponse As Byte
            sMsg = "That Invoice Number was Cancelled Previously. Do you want to use the same invoice number?"
            bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
            If bResponse = vbYes Then
               ' Clear the Invoice number whenthe form got loded.
               sSql = "DELETE FROM CihdTable WHERE INVNO=" & lNewInv _
                      & " AND INVTYPE='TM'"
               clsADOCon.ExecuteSql sSql
               'RdoCon.Execute sSql, rdExecDirect
               lNewInv = txtInv
               success = UpdateInvoice(bAllowReuse)
            End If
            
         Else
            sMsg = "That Invoice Number Is In Use."
            MsgBox sMsg, vbInformation, Caption
            txtInv = Format(lNewInv, "000000")
         End If
      Else
         If iTotalItems = 0 Then
            sMsg = "No Items On This Packing Slip To Invoice."
            MsgBox sMsg, vbInformation, Caption
         Else
            success = UpdateInvoice()
         End If
      End If
   End If
   
    'If Not success Then
    '    MsgBox "Invoice generation failed.  Please try again.", , "ERROR"
    'End If
    
End Sub

Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
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
      SelectHelpTopic Me, "Customer Invoice (Packing Slip)"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdVew_Click()
   Dim i As Integer
   Dim A As Integer
   Dim sQty As String
   Dim sItem As String * 11
   Dim sString As String * 20
   GetPsItems
   If iTotalItems = 0 Then
      MsgBox "No Items On This Packing Slip To Invoice.", vbInformation, Caption
      Exit Sub
   End If
   For i = 1 To iTotalItems
      If vItems(i, 10) Then
         sItem = vItems(i, 0) & "-" & vItems(i, 1) _
                 & vItems(i, 2)
         sString = Left(vItems(i, 3), 20)
         sQty = vItems(i, 4)
         A = Len(sQty)
         VewInvItem.lstItm.AddItem sItem _
            & " " & sString _
            & String(11 - A, Chr(160)) _
            & vItems(i, 4)
      End If
   Next
   VewInvItem.Show 1
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "SJ", ES_SYSDATE, sJournalID
      lblJrn.ForeColor = ES_RED
      lblJrn = "Warning: One Or More Accounts Required."
      
      'Secure.UserInitials = "MGR"
      GetSJAccounts
      CheckInvAsPS
      GetNextInvoice
      FillCombo
      AddInvoice
      GetOptions
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   
   sAccount = GetSetting("Esi2000", "EsiFina", "LastRAccount", sAccount)
   sInvPre = GetSetting("Esi2000", "EsiFina", "LastInvPref", sInvPre)
   If Len(Trim(sInvPre)) = 0 Then sInvPre = "I"
   
   On Error Resume Next
   sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
          & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP,PSPRIMARYSO,CUTYPE FROM PshdTable," _
          & "CustTable,PsitTable WHERE CUREF=PSCUST AND (PSTYPE=1 AND " _
          & "PSSHIPPRINT=1 AND PSINVOICE=0) AND PSNUMBER=PIPACKSLIP AND PSNUMBER= ? "
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 8
   AdoQry1.parameters.Append AdoParameter1
   
   
   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM,PISOREV,PISELLPRICE," & vbCrLf _
          & "PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT,PASTDCOST,SOTAXABLE" & vbCrLf _
          & "FROM PsitTable" & vbCrLf _
          & "JOIN PartTable ON PIPART=PARTREF" & vbCrLf _
          & "JOIN  SohdTable on PISONUMBER = SONUMBER" & vbCrLf _
          & "WHERE PIPACKSLIP= ?" & vbCrLf _
          & "ORDER BY PISONUMBER,PISOITEM"
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adChar
   AdoParameter2.SIZE = 8
   AdoQry2.parameters.Append AdoParameter2
   
   lblTot = "0.00"
   lblItm = "0.00"
   txtTax = "0.00"
   txtFrt = "0.00"
   txtDte = Format(GetServerDateTime, "mm/dd/yy")
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Dim bResponse As Byte
   'Dim sMsg      As String
   'If bCancel And iTotalItems > 0 Then
   '    sMsg = "Do You Really Want To Cancel The Addition" & vbCrLf _
   '        & "Of Invoice " & txtInv & " ?"
   '    bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   '    bCancel = False
   '    If bResponse = vbNo Then Cancel = True
   'End If
   On Error Resume Next
   
   
   Dim RdoInv As ADODB.Recordset
   
   sSql = "SELECT INVCNT FROM CihdTable WHERE INVNO = " & Val(lNewInv) _
               & " AND INVTYPE = 'TM' AND INVCNT  = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   If bSqlRows Then
      'dump any left over dummies
      
      sSql = "DELETE FROM CihdTable WHERE INVNO=" & lNewInv _
             & " AND INVTYPE='TM'"
      clsADOCon.ExecuteSql sSql
      
   Else
      'Update inv count
      sSql = "UPDATE CihdTable SET INVCNT = ISNULL(INVCNT, 1) - 1  " _
                & " WHERE INVNO = " & Val(lNewInv) & " AND INVTYPE = 'TM'"
      clsADOCon.ExecuteSql sSql
   
   End If
   Set RdoInv = Nothing
   
   
   'dump any left over dummies
   'sSql = "DELETE FROM CihdTable WHERE INVNO=" & lNewInv _
   '       & " AND INVTYPE='TM' AND INVUSR = '" & Secure.UserInitials & "'"
   'clsADOCon.ExecuteSQL sSql
   
   SaveSetting "Esi2000", "EsiFina", "LastInvPref", cmbPre
   If Len(Trim(sAccount)) Then SaveSetting "Esi2000", "Fina", "LastRAccount", sAccount
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter1 = Nothing
   Set AdoQry1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoQry2 = Nothing
   Set diaARe02a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(txtDte, "mm/dd/yy")
End Sub

Public Sub FillCombo()
   
   Dim i As Integer
   cmbPre.Clear
   cmbPsl.Clear
   
   On Error GoTo DiaErr1
   
   Dim RdoCmb As ADODB.Recordset
   cmbPre = sInvPre
   For i = 65 To 90
      
      AddComboStr cmbPre.hWnd, Chr$(i)
   Next
   
   sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM PshdTable," _
          & "PsitTable WHERE (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 AND PSSHIPPED=1) " _
          & "AND PSNUMBER=PIPACKSLIP"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbPsl = "" & Trim(!PsNumber)
         Do Until .EOF
            AddComboStr cmbPsl.hWnd, "" & Trim(!PsNumber)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   bGoodPs = GetPackslip()
   GetPsItems
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetPackslip() As Boolean
   Dim RdoPsl As ADODB.Recordset
   
   On Error GoTo DiaErr1
   Erase vItems
   sTaxCode = ""
   sTaxState = ""
   sTaxAccount = ""
   nTaxRate = 0
   AdoQry1.parameters(0).Value = Trim(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoPsl, AdoQry1, ES_DYNAMIC, False, 1)
   If bSqlRows Then
      With RdoPsl
         lblCst.ForeColor = Me.ForeColor
         lblNme.ForeColor = Me.ForeColor
         cmbPsl = "" & Trim(!PsNumber)
         lblCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         sPsCust = "" & Trim(!CUREF)
         sPsStadr = "" & Trim(!PSSTNAME) & vbCrLf _
                    & Trim(!PSSTADR)
         cFREIGHT = Format(!PSFREIGHT, "#####0.00")
         txtFrt = Format(cFREIGHT, "#####0.00")
         lSo = !PSPRIMARYSO
         .Cancel
      End With
      
      GetSalesTaxInfo Compress(lblCst), nTaxRate, sTaxCode, sTaxState, sTaxAccount
      sPsStadr = CheckComments(sPsStadr)
      cmdAdd.enabled = True
      GetPackslip = True
   Else
      
      cFREIGHT = 0
      cmdAdd.enabled = False
      GetPackslip = False
      lblCst.ForeColor = ES_RED
      lblNme.ForeColor = ES_RED
      lblCst = "** Invalid **"
      lblNme = "Packing Slip Invalid Or Hasn't Been Printed."
   End If
   Set RdoPsl = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   lblCst = "*** Invalid ***"
   lblNme = "Packing Slip Invalid Or Hasn't Been Printed."
   DoModuleErrors Me
   
End Function

Public Sub GetNextInvoice()
   Dim inv As New ClassARInvoice
   Dim bDup As Boolean
   lNewInv = inv.GetNextInvoiceNumber
   
   If (chkInvPS.Value <> 1) Then
      txtInv = Format(lNewInv, "000000")
   Else
      ' Disable the 2 invoice sel
      txtInv.enabled = False
      cmbPre.enabled = False
      
      Dim strPSNum As String
      strPSNum = Mid$(CStr(cmbPsl), 3, Len(cmbPsl))
      If (strPSNum <> "") Then
         lNewInv = Val(strPSNum)
         txtInv = Format(strPSNum, "000000")
'      Else
'         MsgBox "PackSlip number is empty.", vbInformation, Caption
'         Exit Sub
      End If
   
   End If

   ' Validate the Invoice number
   If (Trim(txtInv) <> "") Then
      bDup = inv.DuplicateInvNumber(CLng(txtInv))
      
      If (bDup = True) Then
         MsgBox "Invoice number exists.", vbInformation, Caption
         Exit Sub
      End If
   End If
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = CheckComments(txtCmt)
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 8)
   txtFrt = Format(Abs(Val(txtFrt)), "####0.00")
   cFREIGHT = CCur(txtFrt)
   UpdateTotals
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 6)
   txtInv = Format(Abs(Val(txtInv)), "000000")
End Sub

Public Sub AddInvoice()

   Dim RdoInvc As ADODB.Recordset
   GetNextInvoice
   On Error GoTo DiaErr1
   ' Reserve a record in case the invoice number is changed.
   ' Use TM so that it won't show and can be safely deleted.
   
   clsADOCon.BeginTrans
   sSql = "SELECT * FROM CihdTable WHERE INVNO = " & Val(lNewInv) & " AND INVTYPE = 'TM'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInvc, ES_KEYSET)
   If Not bSqlRows Then
      sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVSO,INVCANCELED, INVCNT,INVUSR) " _
             & "VALUES(" & lNewInv & ",'TM'," _
             & Val(cmbPsl) & ",0,1,'" & Secure.UserInitials & "')"
      clsADOCon.ExecuteSql sSql
   Else
      sSql = "UPDATE CihdTable SET INVCNT = ISNULL(INVCNT, 1) + 1  " _
                & " WHERE INVNO = " & Val(lNewInv) & " AND INVTYPE = 'TM'"
      clsADOCon.ExecuteSql sSql
   End If
   clsADOCon.CommitTrans
   
   Set RdoInvc = Nothing
   
   
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Public Function UpdateInvoice(Optional bAllowInvReuse As Boolean = False) As Boolean
    'returns True if successful or canceled
    'returns False if transaction failed
    
   UpdateInvoice = True
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim success As Boolean
   success = True
   
   Dim A As Integer
   Dim i As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   
   Dim nLDollars As Currency
   Dim nTdollars As Currency
   
   Dim sPart As String
   Dim iLevel As Integer
   Dim cCost As Currency
   Dim sProd As String
   
   ' Accounts most of which are not used in this transaction yet.
   Dim sPartRevAcct As String
   Dim sPartCgsAcct As String
   Dim sRevAccount As String
   Dim sDisAccount As String
   Dim sCGSMaterialAccount As String
   Dim sCGSLaborAccount As String
   Dim sCGSExpAccount As String
   Dim sCGSOhAccount As String
   Dim sInvMaterialAccount As String
   Dim sInvLaborAccount As String
   Dim sInvExpAccount As String
   Dim sInvOhAccount As String
   
   ' BnO Taxes
   Dim nRate As Currency
   Dim sType As String
   Dim sState As String
   Dim sCode As String
   
   Dim sPost As String
   
   Dim sTemp As String
   
   ' Make sure user really wants to proceed.
   sMsg = "Post The Selected Invoice With Packing Slip?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Function
   End If
   Dim solist As String
   For i = 1 To iTotalItems
      If solist <> "" Then solist = solist & ","
      solist = solist & vItems(i, 0)
   Next
   
   Dim rs As ADODB.Recordset
   Dim sos As String
   sSql = "select distinct SONUMBER from SohdTable" & vbCrLf _
      & "where SOITAREAR = 1 and SONUMBER in (" & solist & ")"
      
   'sSql = sSql & vbCrLf & "order by SONUMBER"
   If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) <> 0 Then
      With rs
         Do Until rs.EOF
            If Len(sos) > 0 Then
               sos = sos & ","
            End If
            sos = sos & !SONUMBER
            .MoveNext
         Loop
      End With
      MsgBox "SOs with ITAR/EAR status: " & sos
   End If
   rs.Close
   Set rs = Nothing
   

   
   
   
   ' Look For Accounts ?
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   End If
   If iTrans > 0 Then
      bByte = True
      For i = 1 To iTotalItems
         If Val(vItems(i, 8)) > 0 Then
            sPart = vItems(i, 3)
            iLevel = Val(vItems(i, 7))
            sProd = vItems(i, 8)
            bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
                    sPartRevAcct, , sPartCgsAcct)
         End If
         If bByte = False Then Exit For
      Next
   End If
   
   lCurrInvoice = Val(txtInv)
   sPost = Format(txtDte, "mm/dd/yyyy")
   
   cmdAdd.enabled = False
   Prg1.Visible = True
   Prg1.Value = 10
   
   On Error Resume Next
   success = True
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
    ' if allow reuse cancelled invoice
    If (bAllowInvReuse) Then
   
        ' First copy the cancelled date
        If success Then
            sSql = "UPDATE CihdTable SET INVORGCANCDATE = INVCANCDATE " _
                   & " WHERE INVNO=" & lNewInv & " AND INVCANCELED = 1"
            success = clsADOCon.ExecuteSql(sSql)
        End If
        
        ' Get
        If success Then
            sSql = "UPDATE CihdTable SET INVCANCDATE = NULL, INVCANCELED = 0," _
                & "INVTYPE='TM' WHERE INVNO=" & lNewInv & " AND INVCANCELED = 1"
            success = clsADOCon.ExecuteSql(sSql)
        End If
        
    End If
   
    If success Then
        sSql = "UPDATE PshdTable SET PSINVOICE=" & lCurrInvoice & " " _
               & "WHERE PSNUMBER='" & cmbPsl & "' "
        success = clsADOCon.ExecuteSql(sSql)
    End If
        
    ' Update lot record
    If success Then
        sSql = "UPDATE LoitTable SET LOICUSTINVNO=" & lCurrInvoice _
               & ",LOICUST='" & sPsCust & "' WHERE LOIPSNUMBER = '" & cmbPsl & "'"
        success = clsADOCon.ExecuteSql(sSql)
    End If
    
    A = 10
    
    For i = 1 To iTotalItems
      
      ' Progress bar stuff
      A = A + 5
      If A > 80 Then A = 80
      Prg1.Value = A
      
      ' Running invoice total
      nTdollars = nTdollars + (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      
      ' Part accounts
      sPart = vItems(i, 3)
      sProd = vItems(i, 8)
      iLevel = Val(vItems(i, 7))
      bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
              sRevAccount, sDisAccount, sCGSMaterialAccount, sCGSLaborAccount, _
              sCGSExpAccount, sCGSOhAccount, sInvMaterialAccount, sInvLaborAccount, _
              sInvExpAccount, sInvOhAccount)
      
      ' BnO tax
      sCode = ""
      nRate = 0
      sState = ""
      sType = ""
      GetPartBnO vItems(i, 3), nRate, sCode, sState, sType
      If sCode = "" Then
         GetCustBnO sPsCust, nRate, sCode, sState, sType
      End If
      
      ' Update the sales order item
    If success Then
      sSql = "UPDATE SoitTable SET " _
             & "ITINVOICE=" & lCurrInvoice & "," _
             & "ITREVACCT='" & sRevAccount & "'," _
             & "ITCGSACCT='" & sCOSjARAcct & "'," _
             & "ITBOSTATE='" & sState & "'," _
             & "ITBOCODE='" & sCode & "'," _
             & "ITSLSTXACCT='" & sTaxAccount & "'," _
             & "ITTAXCODE='" & sTaxCode & "'," _
             & "ITSTATE='" & sTaxState & "'," _
             & "ITTAXRATE=" & nTaxRate & "," _
             & "ITTAXAMT=" & CCur((nTaxRate / 100) * (Val(vItems(i, 4)) * Val(vItems(i, 11)))) & " " _
             & "WHERE ITSO=" & Val(vItems(i, 0)) & " AND " _
             & "ITNUMBER=" & Val(vItems(i, 1)) & " AND " _
             & "ITREV='" & vItems(i, 2) & "' "
      success = clsADOCon.ExecuteSql(sSql)
    End If
      
      'Journal entries
      nLDollars = (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      cCost = (Val(vItems(i, 4)) * Val(vItems(i, 9)))

      Dim gl As New GLTransaction
      gl.JournalID = Trim(sJournalID)
      gl.InvoiceDate = CDate(txtDte)
      gl.InvoiceNumber = lCurrInvoice
      
        ' Debit A/R (+)
        If success Then
            success = gl.AddDebitCredit(CCur(nLDollars), 0, Compress(sCOSjARAcct), sPart, _
                                 CLng(vItems(i, 0)), CInt(vItems(i, 1)), CStr(vItems(i, 2)), sPsCust, "", True)
        End If
      
      ' Credit Revenue (-)
        If success Then
            success = gl.AddDebitCredit(0, CCur(nLDollars), sRevAccount, _
                              sPart, CLng(vItems(i, 0)), CInt(vItems(i, 1)), CStr(vItems(i, 2)), sPsCust, "", True)
        End If
        Set gl = Nothing
      
    Next
   
   ' Tax and freight
   Prg1.Value = 90
   If cFREIGHT > 0 Then
      
      ' Debit A/R Freight
        If success Then
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
                   & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
                   & "VALUES('" & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cFREIGHT & ",'" _
                   & sCOSjARAcct & "','" _
                   & sPsCust & "','" _
                   & Format(txtDte, "mm/dd/yy") & "'," _
                   & lCurrInvoice & ")"
            success = clsADOCon.ExecuteSql(sSql)
        End If
      
      ' Credit Freight
        If success Then
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
                   & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
                   & "VALUES('" & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cFREIGHT & ",'" _
                   & sCOSjNFRTAcct & "','" _
                   & sPsCust & "','" _
                   & Format(txtDte, "mm/dd/yy") & "'," _
                   & lCurrInvoice & ")"
            success = clsADOCon.ExecuteSql(sSql)
        End If
   End If
   
   cTax = CCur(txtTax)
   If cTax > 0 Then
      ' Debit A/R Taxes
      iRef = iRef + 1
        If success Then
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
                   & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
                   & "VALUES('" & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cTax & ",'" _
                   & sCOSjARAcct & "','" _
                   & sPsCust & "','" _
                   & Format(txtDte, "mm/dd/yy") & "'," _
                   & lCurrInvoice & ")"
            success = clsADOCon.ExecuteSql(sSql)
        End If
      
      ' Credit Taxes
        If success Then
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
                   & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
                   & "VALUES('" & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cTax & ",'" _
                   & sCOSjTaxAcct & "','" _
                   & sPsCust & "','" _
                   & Format(txtDte, "mm/dd/yy") & "'," _
                   & lCurrInvoice & ")"
            success = clsADOCon.ExecuteSql(sSql)
        End If
   End If
   
   ' Change TM invoice to PS
        If success Then
            Prg1.Value = 95
            sSql = "UPDATE CihdTable SET INVNO=" & lCurrInvoice & "," _
                   & "INVPRE='" & cmbPre & "',INVSTADR='" & sPsStadr & "'," _
                   & "INVTYPE='PS',INVSO=0," _
                   & "INVCUST='" & sPsCust & "' WHERE " _
                   & "INVNO=" & lNewInv & " AND INVTYPE='TM'"
            success = clsADOCon.ExecuteSql(sSql)
        End If
   
   If (chkInvPS.Value <> 1) Then
      Dim inv As New ClassARInvoice
      inv.SaveLastInvoiceNumber lNewInv
   End If
   
   ' Add freight and tax to invoice total
   nTdollars = nTdollars + (cTax + cFREIGHT)
   
   ' Then post the total to the invoice
    If success Then
        sSql = "UPDATE CihdTable SET INVTOTAL=" & nTdollars & "," _
               & "INVFREIGHT=" & cFREIGHT & "," _
               & "INVTAX=" & cTax & "," _
               & "INVSHIPDATE='" & sPost & "'," _
               & "INVDATE='" & sPost & "'," _
               & "INVCOMMENTS='" & txtCmt & "'," _
               & "INVPACKSLIP='" & cmbPsl & "' " _
               & "WHERE INVNO=" & lCurrInvoice & " "
        success = clsADOCon.ExecuteSql(sSql)
    End If
   
    If success Then
        sSql = "UPDATE PshdTable SET PSFREIGHT=" & cFREIGHT _
               & " WHERE PSNUMBER='" & cmbPsl & "' "
        success = clsADOCon.ExecuteSql(sSql)
    End If
   
   MouseCursor 0
   
'   If clsADOCon.ADOErrNum = 0 Then
    If success Then
      clsADOCon.CommitTrans
      UpdateInvoice = True
      Prg1.Value = 100
    Else
      clsADOCon.RollbackTrans
      UpdateInvoice = False
      clsADOCon.ADOErrNum = 0
      Prg1.Value = 0
      sMsg = "Could not post the invoice.  Please try again."
      MsgBox sMsg, vbInformation, Caption
      Prg1.Visible = False
      Exit Function
    End If
   
   ' Let the user know the transaction it complete and successful.
   ' Also ask to print.
   Sleep 500
   sMsg = "Invoice " & cmbPre & Format(lCurrInvoice, "000000") _
          & " Was Successfully Posted." & vbCrLf & vbCrLf _
          & "Do You Wish To Print It?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION)
   If bResponse = vbYes Then
      diaARp01a.Visible = False
      Show diaARp01a
      'diaARp01a.cmbInv = lNewInv
      sTemp = diaARp01a.lblPrinter
      diaARp01a.lblPrinter = lblPrinter
      DoEvents    'make sure inv combobox filled before setting inv #
      diaARp01a.cmbInv = lCurrInvoice        'Was lNewInv
      diaARp01a.txtCopies = Me.txtCopies
      diaARp01a.optPrn = True    'click the print button
      diaARp01a.lblPrinter = sTemp
      Unload diaARp01a

'      diaARp01a.Visible = False
'      DoEvents
'      diaARp01a.PrintPS lNewInv
'      Unload diaARp01a

      Me.SetFocus
   End If
   
   ' Reset all values clear status bars and create a new temp invoice.
   Prg1.Visible = False
   iTotalItems = 0
   lblItm = "0.00"
   lblTot = "0.00"
   txtTax = "0.00"
   txtCmt = ""
   ' We don't want to reset the date
   'txtDte = Format(GetServerDateTime, "mm/dd/yy")
   Erase vItems
   'GetNextInvoice
   FillCombo
   AddInvoice
   Exit Function
   
'DiaErr1:   ' no longer used
'   sProcName = "updateinvoice"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
End Function

Public Sub GetPsItems()
   Dim i As Integer
   Dim RdoPsi As ADODB.Recordset
   'Dim cTax As Currency
   Dim cItm As Currency
   Dim cTot As Currency
   
   MouseCursor 13
   On Error Resume Next
   cTax = 0
   
   ' Update the packslip dollars from sales order
   sSql = "UPDATE PsitTable SET PISELLPRICE=" _
          & "ITDOLLARS FROM PsitTable,SoitTable WHERE " _
          & "(PISONUMBER=ITSO AND PISOITEM=ITNUMBER " _
          & "AND PISOREV=ITREV) AND PIPACKSLIP='" & Trim(cmbPsl) & "'"
   clsADOCon.ExecuteSql sSql
   
   On Error GoTo DiaErr1
   
   ' Return packslip items
   AdoQry2.parameters(0).Value = Trim(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoPsi, AdoQry2)
   If bSqlRows Then
      With RdoPsi
         Do Until .EOF
            i = i + 1
            vItems(i, 0) = Format(!PISONUMBER, SO_NUM_FORMAT)
            vItems(i, 1) = Format(!PISOITEM, "##0")
            vItems(i, 2) = Trim(!PISOREV)
            vItems(i, 3) = "" & Trim(!PartRef)
            'vItems(i, 4) = Format(!PIQTY, "#####0.000")
            vItems(i, 4) = !PIQTY
            vItems(i, 5) = "" & sAccount
            vItems(i, 6) = "" & !PIPART
            vItems(i, 7) = "" & !PALEVEL
            vItems(i, 8) = "" & Trim(!PAPRODCODE)
            'vItems(i, 9) = Format(!PASTDCOST, "#####0.000")
            vItems(i, 9) = !PASTDCOST
            vItems(i, 10) = 1
            'vItems(i, 11) = Format(!PISELLPRICE, "#####0.000")
            vItems(i, 11) = !PISELLPRICE
            vItems(i, 12) = !PATAXEXEMPT
            If vItems(i, 12) = 0 And !SOTAXABLE = 1 Then
               cTax = cTax + ((vItems(i, 11) * vItems(i, 4)) * (nTaxRate / 100))
            End If
            cItm = cItm + (vItems(i, 11) * vItems(i, 4))
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoPsi = Nothing
   iTotalItems = i
   lblItm = Format(cItm, CURRENCYMASK)
   txtTax = Format(cTax, CURRENCYMASK)
   UpdateTotals
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getpsitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   lblCst = "** Invalid **"
   lblNme = "Packing Slip Invalid Or Hasn't Been Printed."
   DoModuleErrors Me
End Sub

Private Sub txtTax_LostFocus()
   txtTax = CheckLen(txtTax, 8)
   txtTax = Format(Abs(Val(txtTax)), "0.00")
   UpdateTotals
End Sub

Private Sub UpdateTotals()
   Dim cTot As Currency
   cTot = CCur(lblItm) + CCur(txtTax) + CCur(txtFrt)
   lblTot = Format(cTot, CURRENCYMASK)
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(cmbPre)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions & Left(txtCopies & Space(2), 2)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(GetSetting("Esi2000", "EsiFina", Me.Name, sOptions))
   If sOptions <> "" Then
      cmbPre = Left(sOptions, 1)
      If Len(sOptions) > 1 Then txtCopies = Trim(Mid(sOptions, 2, 2))
   End If
   If Len(Trim(txtCopies)) = 0 Then txtCopies = "1"
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub


Private Function SetInvNumber()
   
   If (chkInvPS.Value <> 1) Then
      GetNextInvoice
   Else
      Dim strPSNum As String
      strPSNum = Mid$(Val(cmbPsl), 3, Len(cmbPsl))
      If (strPSNum <> "") Then
         txtInv = Format(strPSNum, "000000")
      Else
         GetNextInvoice
      End If
   End If
   
End Function
   
Private Function DeleteOldTmpInv(strPreInv As String)
   On Error Resume Next
   Dim RdoInvc As ADODB.Recordset
   
   
   sSql = "SELECT INVCNT FROM CihdTable WHERE INVNO = " & Val(strPreInv) _
               & " AND INVTYPE = 'TM' AND INVCNT  = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInvc, ES_KEYSET)
   If bSqlRows Then
      'dump any left over dummies
      sSql = "DELETE FROM CihdTable WHERE INVNO=" & Val(strPreInv) _
             & " AND INVTYPE='TM' AND INVUSR = '" & Secure.UserInitials & "'"
      clsADOCon.ExecuteSql sSql
   Else
      'Update inv count
      sSql = "UPDATE CihdTable SET INVCNT = ISNULL(INVCNT, 1) - 1  " _
                & " WHERE INVNO = " & Val(strPreInv) & " AND INVTYPE = 'TM'"
      clsADOCon.ExecuteSql sSql
   
   End If
   Set RdoInvc = Nothing
   
   
   
End Function

Private Function CheckInvAsPS()
   Dim RdoInv As ADODB.Recordset
   
   sSql = "SELECT * FROM ComnTable WHERE COALLOWINVNUMPS = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   If bSqlRows Then
      'chkInvPS.enabled = True
      chkInvPS.Value = 1
      ClearResultSet RdoInv
   Else
      'chkInvPS.enabled = False
      chkInvPS.Value = 0
   End If
   Set RdoInv = Nothing
   
End Function

