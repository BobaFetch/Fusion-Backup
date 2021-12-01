VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaMproj
   BorderStyle = 3 'Fixed Dialog
   Caption = "Charge Material To A Project"
   ClientHeight = 3555
   ClientLeft = 1620
   ClientTop = 960
   ClientWidth = 6795
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 3555
   ScaleWidth = 6795
   ShowInTaskbar = 0 'False
   Begin VB.Frame Frame1
      Height = 30
      Left = 0
      TabIndex = 22
      Top = 1560
      Width = 6855
   End
   Begin Threed.SSFrame fra1
      Height = 15
      Left = 120
      TabIndex = 20
      Top = 1500
      Width = 6705
      _Version = 65536
      _ExtentX = 11827
      _ExtentY = -26
      _StockProps = 14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.CommandButton cmdChg
      Caption = "&Add"
      Height = 315
      Left = 5780
      TabIndex = 4
      ToolTipText = "Charge This Item To The Project"
      Top = 1680
      Width = 915
   End
   Begin VB.TextBox txtQty
      Height = 285
      Left = 4560
      TabIndex = 3
      ToolTipText = "Adjustment Quantity"
      Top = 2640
      Width = 1095
   End
   Begin VB.ComboBox cmbPpr
      BackColor = &H00FFFFFF&
      ForeColor = &H00000000&
      Height = 315
      Left = 1200
      TabIndex = 2
      ToolTipText = " Part Number To Be Charged"
      Top = 2280
      Width = 3255
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 1200
      TabIndex = 0
      ToolTipText = "Select Project Part Number"
      Top = 720
      Width = 3545
   End
   Begin VB.ComboBox cmbRun
      ForeColor = &H00800000&
      Height = 315
      Left = 5280
      TabIndex = 1
      ToolTipText = "Select Run Number"
      Top = 720
      Width = 1095
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5780
      TabIndex = 5
      TabStop = 0 'False
      Top = 120
      Width = 915
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 120
      Top = 3000
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3555
      FormDesignWidth = 6795
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 21
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaMproj.frx":0000
      PictureDn = "diaMproj.frx":0146
   End
   Begin VB.Label lblTyp
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1680
      TabIndex = 19
      Top = 3000
      Width = 375
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Type "
      Height = 255
      Index = 7
      Left = 1200
      TabIndex = 18
      Top = 3000
      Width = 495
   End
   Begin VB.Label lblCst
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Left = 4560
      TabIndex = 17
      Top = 3000
      Visible = 0 'False
      Width = 975
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Material"
      Height = 255
      Index = 6
      Left = 120
      TabIndex = 16
      Top = 2280
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Uom     "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 5
      Left = 5760
      TabIndex = 15
      Top = 2040
      Width = 495
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Qoh/Chg Qty         "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 3
      Left = 4560
      TabIndex = 14
      Top = 2040
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Charged Part Number                                                 "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 4
      Left = 1200
      TabIndex = 13
      Top = 2040
      Width = 3015
   End
   Begin VB.Label lblPsc
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1200
      TabIndex = 12
      Top = 2640
      Width = 3015
   End
   Begin VB.Label lblUom
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 5760
      TabIndex = 11
      Top = 2280
      Width = 495
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4560
      TabIndex = 10
      Top = 2280
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Project"
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 9
      Top = 480
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 8
      Top = 720
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Run"
      Height = 255
      Index = 1
      Left = 4800
      TabIndex = 7
      Top = 720
      Width = 615
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1200
      TabIndex = 6
      Top = 1080
      Width = 3255
   End
End
Attribute VB_Name = "diaMproj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaMproj - Charge Material To A Project (Part Type 8)
'
' Created: (cjs)
' Revisions:
'   6/21/02 (nth) Fixed runtime error in fillcombo
'
'*********************************************************************************


Dim RdoQry As rdoQuery
Dim bOnLoad As Byte
Dim bGoodMat As Byte
Dim bGoodRuns As Byte

Dim sPartRef As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
End Sub

Private Sub cmbPpr_Click()
   bGoodMat = FindMatPart()
End Sub

Private Sub cmbPpr_LostFocus()
   cmbPpr = CheckLen(cmbPpr, 30)
   bGoodMat = FindMatPart()
End Sub

Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
End Sub

Private Sub cmbprt_GotFocus()
   cmbPrt_Click
End Sub

Private Sub cmbprt_LostFocus()
   cmbprt = CheckLen(cmbprt, 30)
   bGoodRuns = GetRuns()
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdChg_Click()
   Dim i As Integer
   Dim bResponse As Byte
   Dim sDate As String
   Dim sMsg As String
   Dim sMoRun As String * 9
   Dim sMoPart As String * 31
   Dim sNewPart As String
   
   On Error Resume Next
   sDate = Format(Now, "mm/dd/yy")
   If Val(txtQty) = 0 Then
      MsgBox "You Have Entered a Zero Quantity.", vbInformation, Caption
      On Error Resume Next
      txtQty.SetFocus
      Exit Sub
   Else
      sMsg = "You Have Chosen To Charge " & txtQty & " " & lblUom & vbCr _
             & "Part Number " & cmbPpr & " To The Project." & vbCr _
             & "Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         MouseCursor 13
         On Error GoTo MprojCh1
         sNewPart = Compress(cmbPpr)
         cmdChg.Enabled = False
         i = Len(Trim(Str(cmbRun)))
         i = 5 - i
         sMoPart = cmbprt
         sMoRun = "RUN" & Space$(i) & cmbRun
         GetAccounts sNewPart
         On Error Resume Next
         Err = 0
         RdoCon.BeginTrans
         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                & "PKTYPE,PKPDATE,PKADATE,PKPQTY,PKAQTY) " _
                & "VALUES('" & sNewPart & "','" & sPartRef & "'," _
                & cmbRun & ",10,'" & sDate & "','" & sDate & "'," _
                & txtQty & "," & txtQty & ") "
         RdoCon.Execute sSql, rdExecDirect
         
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & txtQty & " " _
                & "WHERE PARTREF='" & sNewPart & "' "
         RdoCon.Execute sSql, rdExecDirect
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                & "INPQTY,INAQTY,INAMT,INCREDIT,INDEBITACCT,INMOPART,INMORUN) " _
                & "VALUES(10,'" & sNewPart & "'," _
                & "'PICK','" & sMoPart & sMoRun _
                & "',-" & txtQty & ",-" & txtQty & "," & lblCst & "','" _
                & sCreditAcct & "','" & sDebitAcct & "','" & sPartRef & "'," _
                & Val(cmbRun) & ")"
         RdoCon.Execute sSql, rdExecDirect
         MouseCursor 0
         If Err = 0 Then
            RdoCon.CommitTrans
            AverageCost sNewPart
            MsgBox "Material Successfully Charged To Project.", vbInformation, Caption
            txtQty = ""
            lblQty = ""
            lblUom = ""
            lblCst = ""
            On Error Resume Next
            cmbRun.SetFocus
         Else
            RdoCon.RollbackTrans
            sMsg = CurrError.Description & vbCr _
                   & "Could Not Complete Project Charge."
            MsgBox sMsg, vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   Exit Sub
   
   MprojCh1:
   Resume MprojCh2
   CurrError.Description = Err.Description
   MprojCh2:
   MouseCursor 0
   On Error Resume Next
   RdoCon.RollbackTrans
   sMsg = CurrError.Description & vbCr _
          & "Could Not Complete Project Charge."
   MsgBox sMsg, vbExclamation, Caption
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      Dim b As Byte
      ' Make sure we have a open journal for the month
      sJournalID = GetOpenJournal("IJ", Format(Now, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MsgBox "There Is No Open Inventory Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         MouseCursor 0
         Unload Me
         Exit Sub
      End If
      
      FillMaterial
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS<>'CA' OR RUNSTATUS<>'CL')"
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoQry = Nothing
   Set diaMproj = Nothing
End Sub

Public Sub FillCombo()
   Dim RdoPrj As rdoResultset
   Dim b As String
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sProcName = "fillcombo"
   
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE PALEVEL=8 " _
          & "AND RUNREF=PARTREF AND (RUNSTATUS<>'CA' OR RUNSTATUS<>'CL')"
   bSqlRows = GetDataSet(RdoPrj)
   If bSqlRows Then
      With RdoPrj
         cmbprt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PARTNUM) Then
               cmbprt.AddItem "" & Trim(!PARTNUM)
               sTempPart = Trim(!PARTNUM)
            End If
            .MoveNext
         Loop
      End With
      bGoodRuns = GetRuns()
   Else
      MsgBox "No Runs Recorded.", vbExclamation, Caption
   End If
   On Error Resume Next
   Set RdoPrj = Nothing
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetRuns() As Byte
   Dim RdoMat As rdoResultset
   Dim iOldLevel As Integer
   iOldLevel = Val(lblTyp)
   cmbRun.Clear
   sPartRef = Compress(cmbprt)
   
   FindPart Me, cmbprt
   lblTyp = iOldLevel
   On Error GoTo DiaErr1
   RdoQry(0) = sPartRef
   bSqlRows = GetQuerySet(RdoMat, RdoQry)
   If bSqlRows Then
      With RdoMat
         cmbRun = Format(!RUNNO, "####0")
         Do Until .EOF
            cmbRun.AddItem Format(!RUNNO, "####0")
            .MoveNext
         Loop
      End With
      GetRuns = True
   Else
      sPartRef = ""
      GetRuns = False
   End If
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub FillMaterial()
   Dim RdoMat As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL " _
          & "FROM PartTable WHERE (PALEVEL=3 OR PALEVEL=3)"
   bSqlRows = GetDataSet(RdoMat)
   If bSqlRows Then
      With RdoMat
         cmbPpr = Trim(!PARTNUM)
         lblTyp = Format(!PALEVEL, "0")
         Do Until .EOF
            cmbPpr.AddItem "" & Trim(!PARTNUM)
            .MoveNext
         Loop
      End With
      bGoodMat = FindMatPart()
   End If
   Set RdoMat = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "fillmater"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), "####0.000")
End Sub

Public Function FindMatPart() As Byte
   Dim RdoMat As rdoResultset
   Dim sNewPart
   
   sNewPart = Compress(cmbPpr)
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PASTDCOST," _
          & "PAQOH FROM PartTable WHERE (PALEVEL=3 OR PALEVEL=3) " _
          & "AND PARTREF='" & sNewPart & "'"
   bSqlRows = GetDataSet(RdoMat)
   If bSqlRows Then
      With RdoMat
         cmbPpr = "" & Trim(!PARTNUM)
         lblPsc = "" & Trim(!PADESC)
         lblUom = "" & !PAUNITS
         lblCst = Format(!PASTDCOST, "####0.0000")
         lblQty = Format(!PAQOH, "####0.000")
         lblTyp = Format(0 + !PALEVEL, "0")
      End With
      cmdChg.Enabled = True
      FindMatPart = True
   Else
      cmdChg.Enabled = False
      FindMatPart = False
   End If
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "findmatpa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub GetAccounts(sPartNumber As String)
   Dim rdoAct As rdoResultset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & sPartNumber & "'"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         bType = Format(!PALEVEL, "0")
         If bType = 6 Or bType = 7 Then
            sDebitAcct = "" & Trim(!PACGSEXPACCT)
            sCreditAcct = "" & Trim(!PAINVEXPACCT)
         Else
            sDebitAcct = "" & Trim(!PACGSMATACCT)
            sCreditAcct = "" & Trim(!PAINVMATACCT)
         End If
         .Cancel
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If bType = 6 Or bType = 7 Then
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSEXPACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVEXPACCT)
            Else
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSMATACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVMATACCT)
            End If
            .Cancel
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         If bType = 6 Or bType = 7 Then
            sSql = "SELECT COREF,COCGSEXPACCT" & Trim(Str(bType)) & "," _
                   & "COINVEXPACCT" & Trim(Str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         Else
            sSql = "SELECT COREF,COCGSMATACCT" & Trim(Str(bType)) & "," _
                   & "COINVMATACCT" & Trim(Str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         End If
         bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(.rdoColumns(0))
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(.rdoColumns(1))
               .Cancel
            End With
         End If
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   Exit Sub
   
   DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Sub
