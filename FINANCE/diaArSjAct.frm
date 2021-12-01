VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaARCSjAct
   BorderStyle = 3 'Fixed Dialog
   Caption = "Revise Sales Order Account Distributions"
   ClientHeight = 1665
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5370
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 1665
   ScaleWidth = 5370
   ShowInTaskbar = 0 'False
   Begin ComctlLib.ProgressBar ProgressBar1
      Height = 255
      Left = 240
      TabIndex = 4
      Top = 1200
      Width = 4095
      _ExtentX = 7223
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.CommandButton cmdGo
      Caption = "Go"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      Height = 315
      Left = 4440
      TabIndex = 3
      Top = 1200
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4440
      TabIndex = 2
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 0
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaArSjAct.frx":0000
      PictureDn = "diaArSjAct.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4920
      Top = 600
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 1665
      FormDesignWidth = 5370
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "This function scans ALL sales journals validating that the part account shown is the account assigned to the part's product code."
      Height = 765
      Index = 0
      Left = 240
      TabIndex = 1
      Top = 240
      Width = 4065
   End
End
Attribute VB_Name = "diaARCSjAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdGo_Click()
   FixSjAccounts
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
      bOnLoad = False
      
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaARCSjAct = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FixSjAccounts()
   Dim rdoSJ As rdoResultset
   Dim lItems As Long
   
   On Error GoTo 0
   
   MouseCursor 13
   
   'On Error GoTo diaerror1
   
   sSql = "SELECT PartTable.PARTREF,JritTable.DCHEAD,JritTable.DCTRAN,JritTable.DCREF, " _
          & "JritTable.DCACCTNO, PcodTable.PCREVACCT " _
          & "FROM JritTable INNER JOIN " _
          & "JrhdTable ON JritTable.DCHEAD = JrhdTable.MJGLJRNL INNER JOIN " _
          & "PartTable ON JritTable.DCPARTNO = PartTable.PARTREF INNER JOIN " _
          & "PcodTable ON PartTable.PAPRODCODE = PcodTable.PCREF " _
          & "WHERE (JrhdTable.MJTYPE = 'SJ') AND (JritTable.DCCREDIT <> 0)"
   
   bSqlRows = GetDataSet(rdoSJ)
   If bSqlRows Then
      With rdoSJ
         'On Error Resume Next
         While Not .EOF
            If !DCACCTNO <> !PCREVACCT Then
               If lItems = 0 Then RdoCon.BeginTrans
               lItems = lItems + 1
               sSql = "UPDATE JritTable SET DCACCTNO ='" & !PCREVACCT & "'" _
                      & " WHERE DCHEAD='" & !DCHEAD & "' AND DCTRAN = " & !DCTRAN & " AND DCREF = " & !DCREF
               RdoCon.Execute sSql
            End If
            .MoveNext
            Debug.Print lItems
         Wend
      End With
      
      If Err = 0 Then
         If lItems <> 0 Then
            RdoCon.CommitTrans
            MsgBox "Successfully Repaired " & lItems & " Sales Journal Items.", vbInformation, Caption
         Else
            MsgBox "No Incorrect Sales Journal Items Found.", vbInformation, Caption
            
         End If
      Else
         RdoCon.RollbackTrans
         MsgBox "Could Not Repair Sales Journal.", vbExclamation, Caption
      End If
   Else
      MsgBox "No Sales Journal Items Found.", vbInformation, Caption
   End If
   
   Set rdoSJ = Nothing
   MouseCursor 0
   Exit Sub
   
   diaError1:
   
   
End Sub
