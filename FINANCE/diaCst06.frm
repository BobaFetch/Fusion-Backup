VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaCst06
   BorderStyle = 3 'Fixed Dialog
   Caption = "Exploded Proposed Standard Cost Analysis"
   ClientHeight = 4320
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 8205
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4320
   ScaleWidth = 8205
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4800
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4320
      FormDesignWidth = 8205
   End
   Begin VB.CheckBox ChkB
      Caption = "___"
      Height = 255
      Left = 3120
      TabIndex = 26
      Top = 3840
      Width = 855
   End
   Begin VB.CheckBox chkLab
      Caption = "___"
      Height = 255
      Left = 3120
      TabIndex = 25
      Top = 3240
      Width = 855
   End
   Begin VB.CheckBox chkExp
      Caption = "___"
      Height = 255
      Left = 3120
      TabIndex = 24
      Top = 2880
      Width = 855
   End
   Begin VB.CheckBox chkStd
      Caption = "___"
      Height = 255
      Left = 3120
      TabIndex = 23
      Top = 2520
      Width = 975
   End
   Begin VB.CheckBox chkSum
      Caption = "___"
      Height = 255
      Left = 3120
      TabIndex = 22
      Top = 2040
      Width = 855
   End
   Begin VB.ComboBox cmbPrt
      DataSource = "rDt1"
      Height = 315
      Left = 3480
      Sorted = -1 'True
      TabIndex = 15
      Tag = "3"
      ToolTipText = "Enter At Least (1) Leading Character"
      Top = 1080
      Width = 3255
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6960
      TabIndex = 2
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 615
      Left = 6960
      TabIndex = 3
      Top = 360
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaCst06.frx":0000
         Style = 1 'Graphical
         TabIndex = 0
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaCst06.frx":017E
         Style = 1 'Graphical
         TabIndex = 1
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 6
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
      PictureUp = "diaCst06.frx":0308
      PictureDn = "diaCst06.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 7
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaCst06.frx":0594
      PictureDn = "diaCst06.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Treated Like Raw Materials Otherwise)"
      Height = 285
      Index = 12
      Left = 4320
      TabIndex = 21
      Top = 3840
      Width = 3705
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Otherwise, Primary And Secondary Shops Used)"
      Height = 285
      Index = 11
      Left = 4320
      TabIndex = 20
      Top = 3240
      Width = 3585
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Otherwise, Proposed Expense Used)"
      Height = 285
      Index = 10
      Left = 4320
      TabIndex = 19
      Top = 2880
      Width = 3345
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(At All Levels)"
      Height = 285
      Index = 9
      Left = 4320
      TabIndex = 18
      Top = 2520
      Width = 3345
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 285
      Index = 8
      Left = 240
      TabIndex = 17
      Top = 1440
      Width = 1275
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 3480
      TabIndex = 16
      Top = 1440
      Width = 3015
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Update Based On BOM For ""B"" Parts?"
      Height = 285
      Index = 7
      Left = 240
      TabIndex = 14
      Top = 3840
      Width = 3225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Use Labor Cost From Routings?"
      Height = 285
      Index = 6
      Left = 240
      TabIndex = 13
      Top = 3240
      Width = 3225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Use Expense Cost From Routings?"
      Height = 285
      Index = 5
      Left = 240
      TabIndex = 12
      Top = 2880
      Width = 3225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Update Standard Cost?"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 11
      Top = 2520
      Width = 3225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(i. e. Bypass Subassembly Analysis)"
      Height = 285
      Index = 3
      Left = 4320
      TabIndex = 10
      Top = 2040
      Width = 3345
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 2
      Left = 6840
      TabIndex = 9
      Top = 1080
      Width = 1545
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 8
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include Summary Information Only? "
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 5
      Top = 2040
      Width = 3225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Explored Proposed Cost Analysis For Part?"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 4
      Top = 1080
      Width = 3225
   End
End
Attribute VB_Name = "diaCst06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
' diaCst06 - Explode code analysis for part
'
' Created: 11/30/01 (nth)
' Revisions:
'
'
'*********************************************************************************
'
'
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim RdoQry1 As rdoQuery
Dim RdoQry2 As rdoQuery
Dim rdoQry3 As rdoQuery

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart Me
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      MouseCursor 13
      FillCombo
      bOnLoad = False
   End If
   
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   
   ' All cost levels for part
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PALEVEL, PAREVDATE," _
          & "PAEXTDESC, PAMAKEBUY, PALEVLABOR, PALEVEXP, PALEVMATL, PALEVOH," _
          & "PALEVHRS, PASTDCOST, PABOMLABOR, PABOMEXP, PABOMMATL, PABOMOH," _
          & "PABOMHRS, PABOMREV, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL, PATOTOH,PAROUTING,PARRQ " _
          & "FROM PartTable WHERE PARTREF = ?"
   Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   
   ' Part type's 1,2,and 3 with routings
   'sSql = "SELECT PARTREF,OPREF,OPSETUP,OPUNIT,OPUNITHRS,OPSUHRS,WCNOHFIXED,WCNOHPCT," _
   '    & "WCNSTDRATE,WCNSUHRS,WCNUNITHRS FROM WcntTable " _
   '    & "INNER JOIN RtopTable ON WcntTable.WCNREF = RtopTable.OPCENTER " _
   '    & "INNER JOIN RtopTable ON ShopTable.SHPREF = RtopTable.OPSHOP " _
   '    & "LEFT OUTER JOIN PartTable ON RtopTable.OPREF = PartTable.PAROUTING"
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   sSql = "SELECT OPSUHRS,OPUNITHRS,WCNSTDRATE,WCNOHFIXED,SHPRATE,SHPOHTOTAL " _
          & "From RnopTable " _
          & "INNER JOIN WcntTable ON RnopTable.OPCENTER = WcntTable.WCNREF " _
          & "INNER JOIN ShopTable ON RnopTable.OPSHOP = ShopTable.SHPREF " _
          & "WHERE (RnopTable.OPREF = ?)"
   Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   
   
   ' Part type 4's or purchased parts
   
   
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaCst06 = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo diaErr1
   SetMdiReportsize MdiSect
   MouseCursor 0
   Exit Sub
   
   ' Error handeling
   diaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCombo()
   Dim rdoPart As rdoResultset
   
   sSql = "Qry_FillParts"
   bSqlRows = GetDataSet(rdoPart)
   
   If bSqlRows Then
      With rdoPart
         While Not .EOF
            AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoPart = Nothing
   cmbPrt.ListIndex = 0
   Exit Sub
   
   diaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetNextBillLevel(sUsedOnPart As String, sRev As String, _
                            bLevel As Byte)
   
   Dim RdoBom As rdoResultset
   Dim RdoPrt As rdoResultset
   Dim RdoRtn As rdoResultset
   Dim iHrs As Integer
   Dim cLab As Currency
   Dim cMat As Currency
   Dim cExp As Currency
   Dim cOh As Currency
   Dim sParent As String
   Dim dRRQ As Double
   
   On Error GoTo 0
   
   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART = '" & sUsedOnPart _
          & "' AND BMREV = '" & sRev & "'"
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   
   If bSqlRows Then
      With RdoBom
         sParent = !BMASSYPART
         
         Do Until .EOF
            If Err > 0 Or bLevel > 10 Then Exit Do
            
            GetNextBillLevel Trim(!BMPARTREF), !BMREV, bLevel + 1
            RdoQry1(0) = !BMPARTREF
            bSqlRows = GetQuerySet(RdoPrt, RdoQry1, ES_FORWARD)
            
            ' RRQ
            If RdoPrt!PARRQ = 0 Then
               dRRQ = 1
            Else
               dRRQ = RdoPrt!PARRQ
            End If
            
            ' Check if the part has a routing.
            If Len(Trim(RdoPrt!PAROUNTING)) Then
               RdoQry2(0) = RdoPrt!PAROUNTING
               bSqlRows = GetQuerySet(RdoRtn, RdoQry2, ES_FORWARD)
               
               ' Loop through all op's
               If bSqlRows Then
                  Dim dHours As Double
                  Dim dDollars As Double
                  Dim dOhDollars As Double
                  
                  While Not RdoRtn.EOF
                     dHours = RdoRtn!OPSUHRS
                     dHours = (dHours / dRRQ)
                     dHours = dHours + RdoRtn!OPUNITHRS
                     
                     ' Check work work center assignment
                     If RdoRtn!WCSTDRATE > 0 Then
                        dDollars = RdoRtn!WCSTDRATE * dHours
                        If RdoRtn!WCNOHFIXED > 0 Then
                           dOhDollars = RdoRtn!WCNOHFIXED * dHours
                        End If
                        ' No work center assigned
                     Else
                        dDollars = RdoRtn!SHPRATE * dHours
                        dOhDollars = RdoRtn!SHOHTOTAL * dHours
                     End If
                     
                     RdoRtn.MoveNext
                  Wend
               End If
               
               ' No routing for this part
            Else
               iHrs = iHrs + (RdoPrt!PALEVHRS)
               cLab = cLab + (RdoPrt!PALEVLABOR)
               cMat = cMat + (RdoPrt!PALEVMATL)
               cExp = cExp + (RdoPrt!PALEVEXP)
               cOh = cOh + (RdoPrt!PALEVOH)
            End If
            .MoveNext
            Set RdoQry1 = Nothing
         Loop
         
         Set RdoPrt = Nothing
         
         sSql = "UPDATE PartTable SET " _
                & "PABOMHRS = " & iHrs _
                & ",PABOMLABOR = " & cLab _
                & ",PABOMMATL = " & cMat _
                & ",PABOMEXP = " & cExp _
                & ",PABOMOH = " & cOh _
                & " WHERE PARTREF = '" & sParent & "'"
         RdoCon.Execute sSql
      End With
   End If
   
   Set RdoBom = Nothing
End Sub


Private Sub optDis_Click()
   UpdateSTDCost
   
   
End Sub


Private Sub UpdateSTDCost()
   Dim sPart As String
   Dim rdoPrt4 As rdoResultset
   Dim rdoPrt123 As rdoResultset
   
   MouseCursor 13
   sPart = Trim(cmbPrt)
   
   RdoCon.BeginTrans
   GetNextBillLevel sPart, "", 1
   If Err = 0 Then
      RdoCon.CommitTrans
   Else
      RdoCon.RollbackTrans
   End If
   MouseCursor 0
   
   
End Sub
