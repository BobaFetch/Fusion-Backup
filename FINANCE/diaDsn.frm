VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaDsn
   BorderStyle = 3 'Fixed Dialog
   Caption = "ODBC Data Source Name"
   ClientHeight = 2040
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 4755
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2040
   ScaleWidth = 4755
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4080
      Top = 720
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2040
      FormDesignWidth = 4755
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 3840
      TabIndex = 3
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin VB.ComboBox cmbDsn
      Height = 315
      Left = 720
      TabIndex = 2
      Top = 960
      Width = 3015
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
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaDsn.frx":0000
      PictureDn = "diaDsn.frx":0146
   End
   Begin VB.Label lblDsn
      BackStyle = 0 'Transparent
      Caption = "Can't Locate Any Installed DataSources."
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00000080&
      Height = 735
      Left = 720
      TabIndex = 5
      Top = 1320
      Width = 4095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Using the Incorrect ODBC Data Source May Prevent Logon To Reports"
      Height = 495
      Index = 2
      Left = 765
      TabIndex = 4
      Top = 480
      Width = 3015
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "DSN"
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 1
      Top = 1005
      Width = 615
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Caution: Make Certain That The DSN Is Correct."
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 0
      Top = 260
      Width = 3615
   End
End
Attribute VB_Name = "diaDsn"
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

Private Sub cmbDsn_GotFocus()
   SelectFormat Me
   
End Sub

Private Sub cmbDsn_LostFocus()
   SaveSetting "Esi2000", "System", "SqlDSN", Trim(cmbDsn)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "ODBC Datasource Name"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   'Move diaSec.Left + 400, diaSec.Top + 700
   cmbDsn.ToolTipText = "Contains A List of SQL Server Datasources." _
                        & "Choose From The List Or Use The Default ESI2000."
   
   sDsn = GetSetting("Esi2000", "System", "SqlDSN", sDsn)
   GetDataSources
   If cmbDsn.ListCount = 0 Then
      lblDsn = lblDsn & vbCr & "ESI2000 Is Default And May Not Be Installed."
      lblDsn.Visible = True
      Height = 2415
   Else
      lblDsn.Visible = False
      Height = 1995
   End If
   cmbDsn = sDsn
   
End Sub




Private Sub Form_Unload(Cancel As Integer)
   Set diaDsn = Nothing
   
End Sub



Public Sub GetDataSources()
   Dim b As Byte
   Dim i As Integer
   Dim iFreeFile As Integer
   Dim sWindows As String
   Dim sEntry As String
   Dim sSqlSvr As String
   
   On Error GoTo DiaErr1
   iFreeFile = FreeFile
   sWindows = GetWindowsDir() & "\ODBC.INI"
   Open sWindows For Input Access Read As #iFreeFile
   i = -1
   Do Until EOF(iFreeFile)
      Input #iFreeFile, sEntry
      If Left(sEntry, 8) = "[ODBC 32" Then
         For b = 1 To 50
            Input #iFreeFile, sEntry
            If Left(sEntry, 1) = "[" Then Exit Do
            sSqlSvr = GetSqlServer(sEntry)
            If sSqlSvr <> "" Then cmbDsn.AddItem sSqlSvr
         Next
      End If
   Loop
   b = 0
   If cmbDsn.ListCount > 0 Then
      For i = 0 To cmbDsn.ListCount - 1
         If cmbDsn.List(i) = "ESI2000" Then b = 1
      Next
   Else
      b = 0
   End If
   If b = 0 Then
      If sServer <> "" Then
         sDsn = RegisterSqlDsn(sDsn)
         cmbDsn.AddItem "ESI2000"
         If cmbDsn = "" Then cmbDsn = cmbDsn.List(0)
      End If
   End If
   Close iFreeFile
   Exit Sub
   
   DiaErr1:
   Close iFreeFile
   On Error GoTo 0
   
End Sub

Public Function GetSqlServer(sDriver As String) As String
   Dim a As Integer
   Dim i As Integer
   Dim sNewEntry As String
   
   On Error Resume Next
   sDriver = Trim$(sDriver)
   a = Len(Trim$(sDriver))
   i = InStr(sDriver, "=")
   If i > 0 Then
      sNewEntry = Mid$(sDriver, i + 1, a - i + 1)
      If UCase$(Left$(sNewEntry, 6)) = "SQL SE" Then
         GetSqlServer = Left(sDriver, i - 1)
      Else
         GetSqlServer = ""
      End If
   Else
      GetSqlServer = ""
   End If
   
End Function



Private Sub lblDsn_Click()
   
End Sub
