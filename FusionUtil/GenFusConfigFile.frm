VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GenFusConfigFile 
   Caption         =   "Generate Fusion Config File"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSrvName 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "SQL Server Administrator"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdDCrpt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   5400
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDNS 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "SQL Server Administrator"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdEncpt 
      Caption         =   "encrypt"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Select Config File"
      Top             =   3960
      Width           =   4455
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   875
   End
   Begin VB.CommandButton cmdGenConfig 
      Caption         =   "&GenerateConfig File"
      Height          =   435
      Left            =   3000
      TabIndex        =   10
      ToolTipText     =   "Update for this user only"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtDBName 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "SQL Server Administrator"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtEncrptPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Verify SQL Server Administrator's Password (15 Char max)"
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox txtPsw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "SQL Server Administrator's Password (15 Char max)"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtLog 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "SQL Server Administrator"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Database Name"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "DNS Name (ODBC)"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Config File Name"
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server Name"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted Password"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Administrator Password"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Administrator Logon"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "GenFusConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDCrpt_Click()
   Dim sEncrypted As String
   Dim sDcrypt As String
   
   sEncrypted = txtEncrptPass.Text
   sDcrypt = GetSecPassword(sEncrypted)
   
   txtEncrptPass.Text = sDcrypt
End Sub

Private Sub cmdEncpt_Click()

   Dim sEncrypted As String, sPassword As String
   sPassword = txtPsw.Text
   sEncrypted = ScramblePw(sPassword)
   txtEncrptPass.Text = sEncrypted
   
End Sub


Public Function ScramblePw(PassW As String) As String
   Dim A As Integer
   Dim b As Integer
   Dim iList As Integer
   Dim NewPassWord As String * 40
   Randomize
   
   On Error GoTo ModErr1
   'First we fill it
   For b = 1 To 40
      iList = Int((48 * Rnd) + 74)
      Mid$(NewPassWord, b, 1) = Chr$(iList)
   Next
   'Now we insert the Password"
   b = Len(PassW)
   If b = 0 Then
      MsgBox "Illegal Password.", vbExclamation, sSysCaption
      Exit Function
   End If
   
   For iList = 3 To 38 Step 2
      'If A = b Then Exit For      'added to allow blank passwords @@@ tel 8/29/07
      A = A + 1
      '1 extra byte
      Mid(NewPassWord, iList, 1) = Chr$(Asc(Mid$(PassW, A, 1)) + 1)
      If A = b Then Exit For
   Next
   'Starts at 3. Where does it end (+2)
   NewPassWord = Left$(NewPassWord, 38) & Format$(iList + 2, "00")
   
   'Turn it around
   For iList = 40 To 1 Step -1
      ScramblePw = ScramblePw & Mid$(NewPassWord, iList, 1)
   Next
   Exit Function
   
ModErr1:
   MsgBox "Illegal Password.", vbExclamation, sSysCaption
   ScramblePw = ""
   
End Function


Public Function GetSecPassword(PassWord As String) As String
   'Unscramble password created from ScramblePw
   Dim A As Integer
   Dim b As Integer
   Dim C As Integer
   Dim TempPw As String
   
   On Error GoTo ModErr1
   
   'Turn it around
   For A = 40 To 1 Step -1
      TempPw = TempPw & Mid$(PassWord, A, 1)
   Next
   C = Val(Right(TempPw, 2)) - 2
   For A = 3 To C Step 2
      GetSecPassword = GetSecPassword & Chr$(Asc(Mid$(TempPw, A, 1)) - 1)
   Next
   Exit Function
   
ModErr1:
   GetSecPassword = ""
   
End Function


Private Sub cmdGenConfig_Click()

   Dim strFullpath As String
   Dim nFileNum As Integer
   'strFilePath = "C:\Development\FusionCode\EDIFiles\Testing\"
   'strFileName = "INVOUT.EDI"
   
   strFullpath = txtFilePath.Text
   
   If (Trim(strFullpath) <> "") Then
      ' Open the file
      nFileNum = FreeFile
      Open strFullpath For Output As nFileNum
      
      If EOF(nFileNum) Then
         GenerateCfgFile nFileNum
      End If
      ' Close the file
      Close nFileNum
   End If

End Sub

Private Function GenerateCfgFile(nFileNum As Integer)
   
      
   If EOF(nFileNum) Then
      
      Dim strKeySec As String
      Dim strSrvName As String
      Dim strSaveDB As String
      Dim strDefDB As String
      Dim strDBName As String
      Dim strSQLUser As String
      Dim strSQLPw As String
      Dim strSQLDSN As String
      
      strSrvName = txtSrvName.Text
      strDBName = txtDBName.Text
      strSQLPw = txtEncrptPass.Text
      strSQLUser = txtLog.Text
      strSQLDSN = txtDNS.Text
   
      strKeySec = "[FUSION_USERSETTINGS]"
      Print #nFileNum, strKeySec
      
      strSaveDB = "SAVEDATABASE = 1"
      Print #nFileNum, strSaveDB
      
      strDefDB = "DEFAULTDATABASE = " & strDBName
      Print #nFileNum, strDefDB
      
      strSrvName = "SERVERNAME = " & strSrvName
      Print #nFileNum, strSrvName
      
      strDBName = "DatabaseName = " & strDBName
      Print #nFileNum, strDBName
      
      strSQLPw = "SQLPASSWORD = " & strSQLPw
      Print #nFileNum, strSQLPw
      
      strSQLUser = "SQLLOGIN = " & strSQLUser
      Print #nFileNum, strSQLUser
      
      strSQLDSN = "SQLDSN = " & strSQLDSN
      Print #nFileNum, strSQLDSN
      
   End If
   
   ' Close the file
   Close nFileNum
End Function

Private Sub cmdOpenDia_Click()
    fileDlg.Filter = "INI File (*.ini) | *.ini"
    fileDlg.ShowOpen
    If fileDlg.FileName = "" Then
        txtFilePath.Text = ""
    Else
        txtFilePath.Text = fileDlg.FileName
    End If
End Sub
