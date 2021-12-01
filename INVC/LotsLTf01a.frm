VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lot Organization - Single Part Number"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "The Expected Actual Quantity On Hand"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtlot 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "User Lot (40)"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdOrg 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      TabIndex        =   15
      ToolTipText     =   "Reorganize Lots And Inventory"
      Top             =   540
      Width           =   875
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Includes Lot Tracked Parts"
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2955
      FormDesignWidth =   8280
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clears Existing Lots (Sets To Zero) And Creates One New Lot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Quantity On Hand"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Lot Number (User ID)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7440
      TabIndex        =   14
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prod Code"
      Height          =   255
      Index           =   18
      Left            =   6120
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   1320
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty On Hand"
      Height          =   255
      Index           =   17
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblQoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7200
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label txtLqoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots QOH"
      Height          =   255
      Index           =   20
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Total Of Lots"
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "LotsLTf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 6/30/03
Option Explicit
Dim bOnLoad As Byte
Dim bGoodPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   bGoodPart = GetPart()
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5550
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdOrg_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If Len(Trim(txtlot)) < 5 Then
      sMsg = "The User Lot Must Be At Leat (5) Characters."
      MsgBox sMsg, vbInformation, Caption
      Beep
      txtlot = "REORG-" & Compress(cmbPrt) & "-" _
               & Format(ES_SYSDATE, "mm/dd/yy")
      Exit Sub
   End If
   If Val(txtQty) = 0 Then
      sMsg = "The Quantity Is Set To Zero." & vbCr _
             & "Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   sMsg = "You Have Chosen To Set All Current Lots To Zero" & vbCr _
          & "And Create One New Lot With A Quantity Of " & txtQty & ". " & vbCr _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      OrganizeLots
      
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   z1(5).ForeColor = ES_BLUE
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PALOTTRACK=1 AND PAINACTIVE = 0 AND PAOBSOLETE = 0   ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart()
   Dim RdoGet As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAQOH," _
          & "PAPRODCODE,PALOTQTYREMAINING FROM PartTable " _
          & "WHERE PARTREF='" & Compress(cmbPrt) & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         GetPart = True
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblDsc.ForeColor = Es_TextForeColor
         lblLvl = "" & Format(!PALEVEL, "0")
         lblQoh = "" & Format(!PAQOH, ES_QuantityDataFormat)
         lblCode = "" & Trim(!PAPRODCODE)
         txtLqoh = "" & Format(!PALOTQTYREMAINING, ES_QuantityDataFormat)
         txtlot = "REORG-" & Trim(!PartRef) & "-" & Format(ES_SYSDATE, "mm/dd/yy")
         txtQty = "0.000"
         ClearResultSet RdoGet
         cmdOrg.Enabled = True
      End With
   Else
      lblDsc = "*** Part Wasn't Found ***"
      lblDsc.ForeColor = ES_RED
      lblLvl = "0"
      lblQoh = "0.000"
      lblCode = ""
      txtLqoh = "0.000)"
      txtlot = ""
      txtQty = "0.000"
      MsgBox "That Part Number Wasn't Found.", vbInformation, _
         Caption
      cmdOrg.Enabled = False
   End If
   Set RdoGet = Nothing
   Exit Function
   
End Function

Private Sub txtlot_LostFocus()
   txtlot = CheckLen(txtlot, 40)
   If lblDsc.ForeColor = ES_RED Then
      txtlot = ""
   Else
      If txtlot = "" Then
         Beep
         txtlot = "REORG-" & Compress(cmbPrt) & "-" _
                  & Format(ES_SYSDATE, "mm/dd/yy")
      End If
   End If
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub



Private Sub OrganizeLots()
   
   Dim lot As New ClassLot
   If lot.ConsolidateLots(cmbPrt.Text, Val(txtQty.Text), txtlot.Text) Then
      SysMsg "Transaction Was Completed", True
   Else
      MsgBox "Could Not Complete The Transaction.  Error " & Err.Description, _
         vbExclamation, Caption
   End If
   cmdOrg.Enabled = True
   Sleep 1000
   bGoodPart = GetPart()
   Exit Sub
   
   
   
   #If False Then
   Dim RdoLot As ADODB.Recordset
   
   Dim iList As Integer
   Dim iLots As Integer
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   
   Dim cActQty As Currency
   Dim cInvQty As Currency
   Dim cLotQty As Currency
   Dim cItmQty As Currency
   
   Dim sLotNum As String
   Dim sPartNumber As String
   Dim vAdate As Variant
   
   Dim sLots(1000) As String
   On Error GoTo DiaErr1
   'Collect the lots
   MouseCursor 11
   Erase sLots
   cmdOrg.Enabled = False
   iLots = 0
   cActQty = Val(txtQty)
   sPartNumber = Compress(cmbPrt)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   'Get the Lots for this part
   sSql = "SELECT LOTNUMBER,LOTPARTREF FROM LohdTable " _
          & "WHERE LOTPARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            iLots = iLots + 1
            sLots(iLots) = "" & Trim(!lotNumber)
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   
   'If we have some then process them
   clsADOCon.ADOErrNum = 0
   RdoCon.BeginTrans
   On Error Resume Next
   For iList = 1 To iLots
      sSql = "SELECT SUM(LOIQUANTITY) FROM LoitTable " _
             & "WHERE LOINUMBER='" & sLots(iList) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
      If bSqlRows Then
         With RdoLot
            Do Until .EOF
               If Not IsNull(.Fields(0)) Then
                  cLotQty = .Fields(0)
               Else
                  cLotQty = 0
               End If
               If cLotQty <> 0 Then
                  sLotNum = sLots(iList)
                  lCOUNTER = GetLastActivity() + 1
                  lLOTRECORD = GetNextLotRecord(sLotNum)
                  If cLotQty > 0 Then
                     sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                            & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                            & "LOIACTIVITY,LOICOMMENT) " _
                            & "VALUES('" _
                            & sLotNum & "'," & lLOTRECORD & ",19,'" & sPartNumber _
                            & "','" & vAdate & "',-" & cLotQty _
                            & "," & lCOUNTER & ",'" _
                            & "Manual Inventory Adjustment" & "')"
                     clsADOCon.ExecuteSQL sSql
                  Else
                     sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                            & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                            & "LOIACTIVITY,LOICOMMENT) " _
                            & "VALUES('" _
                            & sLotNum & "'," & lLOTRECORD & ",19,'" & sPartNumber _
                            & "','" & vAdate & "'," & Abs(cLotQty) _
                            & "," & lCOUNTER & ",'" _
                            & "Manual Inventory Adjustment" & "')"
                     clsADOCon.ExecuteSQL sSql
                  End If
                  
               End If
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=0," _
                      & "LOTAVAILABLE=0 WHERE LOTNUMBER='" & sLots(iList) & "'"
               clsADOCon.ExecuteSQL sSql
               'MsgBox iList
               .MoveNext
            Loop
            ClearResultSet RdoLot
         End With
      End If
   Next
   Set RdoLot = Nothing
   'Processed the lot items.  Now the Part and associated rows
   sSql = "UPDATE PartTable SET PAQOH=" & cActQty & "," _
          & "PALOTQTYREMAINING=" & cActQty & " WHERE " _
          & "PARTREF='" & sPartNumber & "'"
   clsADOCon.ExecuteSQL sSql
   
   sLotNum = GetNextLotNumber()
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
          & "VALUES('" _
          & sLotNum & "','" & txtlot & "','" & sPartNumber _
          & "','" & vAdate & "'," & cActQty & "," & cActQty _
          & ",0,'" & vAdate & "','Manual Lot Re-Org')"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
          & "LOIACTIVITY,LOICOMMENT) " _
          & "VALUES('" _
          & sLotNum & "',1,19,'" & sPartNumber _
          & "','" & vAdate & "'," & cActQty _
          & "," & lCOUNTER & ",'" _
          & "Manual Manual Re-org" & "')"
   clsADOCon.ExecuteSQL sSql
   
   'Now to square Activity away
   sSql = "SELECT SUM(INAQTY) FROM InvaTable WHERE INPART='" _
          & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         If Not IsNull(.Fields(0)) Then
            cInvQty = .Fields(0)
         Else
            cInvQty = 0
         End If
         lCOUNTER = GetLastActivity() + 1
         If cInvQty < 0 Then
            'less than zero
            cInvQty = Abs(cInvQty) + cActQty
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                   & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                   & "VALUES(19,'" & sPartNumber & "','Manual Adjust Reorg','" & txtlot & "'," _
                   & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
                   & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
         Else
            If cInvQty = 0 Then
               cInvQty = cActQty
               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                      & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                      & "VALUES(19,'" & sPartNumber & "','Manual Adj Reorg','" & txtlot & "'," _
                      & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
                      & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
               clsADOCon.ExecuteSQL sSql
            Else
               'Greater than zero
               cInvQty = cInvQty - cActQty
               cInvQty = cInvQty - (2 * cInvQty)
               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                      & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                      & "VALUES(19,'" & sPartNumber & "','Manual Adj Reorg','" & txtlot & "'," _
                      & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
                      & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End With
   End If
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      UpdateWipColumns lCOUNTER
      SysMsg "Transaction Was Completed", True
   Else
      MsgBox Err.Description
      MsgBox "Could Not Complete The Transaction.", _
         vbExclamation, Caption
      clsADOCon.RollbackTrans
   End If
   cmdOrg.Enabled = True
   Sleep 1000
   Set RdoLot = Nothing
   bGoodPart = GetPart()
   Exit Sub
   
DiaErr1:
   sProcName = "organizelots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   #End If
   
End Sub
