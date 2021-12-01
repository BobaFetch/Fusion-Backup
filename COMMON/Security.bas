Attribute VB_Name = "Security"
'Public sSysCaption   As String
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/6/07 Revisited paramenter and made some changes (INVC)
Option Explicit

'Security
'The bytes are 0 or 1 for T/F. Cheaper than boolean

'The following gets the record number for the Current User
'i = GetSetting("Esi2000", "System", "UserProfileRec", i)

Public bRetry As Byte
Public bSecSet As Byte
Public bIdeRunning As Byte

' authorization for user to see various modules
Public bSections(10) As Byte

Public Enum eSection
   SECTION_ADMN = 1
   SECTION_PROD = 2
   SECTION_ENGR = 3
   SECTION_SALE = 4
   SECTION_FINA = 5
   SECTION_QUAL = 6
   SECTION_INVC = 7
   SECTION_TIME = 8
   SECTION_POM = 9
   SECTION_DATACOL = 10

End Enum

Type PassWordClass
   PassWord As String * 40
   UserLcName As String * 30
   UserUcName As String * 30
   UserAdmn As Byte
   UserRecord As Integer
   UserSpare As String * 30
End Type

Public SecPw As PassWordClass

Public Const FIRSTUSERRECORDNO = 3 'first record # in .edd file that may be edited

Type SecurityArray   'alternate structure of .edd record
   UserName As String * 40
   UserInitials As String * 3
   UserNickName As String * 20
   UserActive As Byte
   UserAddUser As Byte
   UserLevel As Byte
   UserNew As String * 8 'Date actually as 10/31/00
   
   UserPermissions(1 To 280) As Byte
End Type

Type SecurityClass 'structure of .edd record
   UserName As String * 40
   UserInitials As String * 3
   UserNickName As String * 20
   UserActive As Byte
   UserAddUser As Byte
   UserLevel As Byte
   
   UserNew As String * 8 'Date actually as 10/31/00
   
   UserAdmn As Byte
   UserAdmnG1 As Byte
   UserAdmnG1E As Byte
   UserAdmnG1V As Byte
   UserAdmnG1F As Byte
   UserAdmnG2 As Byte
   UserAdmnG2E As Byte
   UserAdmnG2V As Byte
   UserAdmnG2F As Byte
   UserAdmnG3 As Byte
   UserAdmnG3E As Byte
   UserAdmnG3V As Byte
   UserAdmnG3F As Byte
   UserAdmnG4 As Byte
   UserAdmnG4E As Byte
   UserAdmnG4V As Byte
   UserAdmnG4F As Byte
   UserAdmnG5 As Byte
   UserAdmnG5E As Byte
   UserAdmnG5V As Byte
   UserAdmnG5F As Byte
   UserAdmnG6 As Byte
   UserAdmnG6E As Byte
   UserAdmnG6V As Byte
   UserAdmnG6F As Byte
   UserAdmnG7 As Byte
   UserAdmnG7E As Byte
   UserAdmnG7V As Byte
   UserAdmnG7F As Byte
   UserAdmnG8 As Byte
   UserAdmnG8E As Byte
   UserAdmnG8V As Byte
   UserAdmnG8F As Byte
   
   UserEngr As Byte
   UserEngrG1 As Byte
   UserEngrG1E As Byte
   UserEngrG1V As Byte
   UserEngrG1F As Byte
   UserEngrG2 As Byte
   UserEngrG2E As Byte
   UserEngrG2V As Byte
   UserEngrG2F As Byte
   UserEngrG3 As Byte
   UserEngrG3E As Byte
   UserEngrG3V As Byte
   UserEngrG3F As Byte
   UserEngrG4 As Byte
   UserEngrG4E As Byte
   UserEngrG4V As Byte
   UserEngrG4F As Byte
   UserEngrG5 As Byte
   UserEngrG5E As Byte
   UserEngrG5V As Byte
   UserEngrG5F As Byte
   UserEngrG6 As Byte
   UserEngrG6E As Byte
   UserEngrG6V As Byte
   UserEngrG6F As Byte
   UserEngrG7 As Byte
   UserEngrG7E As Byte
   UserEngrG7V As Byte
   UserEngrG7F As Byte
   UserEngrG8 As Byte
   UserEngrG8E As Byte
   UserEngrG8V As Byte
   UserEngrG8F As Byte
   
   UserFina As Byte
   UserFinaG1 As Byte
   UserFinaG1E As Byte
   UserFinaG1V As Byte
   UserFinaG1F As Byte
   UserFinaG2 As Byte
   UserFinaG2E As Byte
   UserFinaG2V As Byte
   UserFinaG2F As Byte
   UserFinaG3 As Byte
   UserFinaG3E As Byte
   UserFinaG3V As Byte
   UserFinaG3F As Byte
   UserFinaG4 As Byte
   UserFinaG4E As Byte
   UserFinaG4V As Byte
   UserFinaG4F As Byte
   UserFinaG5 As Byte
   UserFinaG5E As Byte
   UserFinaG5V As Byte
   UserFinaG5F As Byte
   UserFinaG6 As Byte
   UserFinaG6E As Byte
   UserFinaG6V As Byte
   UserFinaG6F As Byte
   UserFinaG7 As Byte
   UserFinaG7E As Byte
   UserFinaG7V As Byte
   UserFinaG7F As Byte
   UserFinaG8 As Byte
   UserFinaG8E As Byte
   UserFinaG8V As Byte
   UserFinaG8F As Byte
   
   'Added just in case
   UserFinaG9 As Byte
   UserFinaG9E As Byte
   UserFinaG9V As Byte
   UserFinaG9F As Byte
   UserFinaG10 As Byte
   UserFinaG10E As Byte
   UserFinaG10V As Byte
   UserFinaG10F As Byte
   
   UserInvc As Byte
   UserInvcG1 As Byte
   UserInvcG1E As Byte
   UserInvcG1V As Byte
   UserInvcG1F As Byte
   UserInvcG2 As Byte
   UserInvcG2E As Byte
   UserInvcG2V As Byte
   UserInvcG2F As Byte
   UserInvcG3 As Byte
   UserInvcG3E As Byte
   UserInvcG3V As Byte
   UserInvcG3F As Byte
   UserInvcG4 As Byte
   UserInvcG4E As Byte
   UserInvcG4V As Byte
   UserInvcG4F As Byte
   UserInvcG5 As Byte
   UserInvcG5E As Byte
   UserInvcG5V As Byte
   UserInvcG5F As Byte
   UserInvcG6 As Byte
   UserInvcG6E As Byte
   UserInvcG6V As Byte
   UserInvcG6F As Byte
   UserInvcG7 As Byte
   UserInvcG7E As Byte
   UserInvcG7V As Byte
   UserInvcG7F As Byte
   UserInvcG8 As Byte
   UserInvcG8E As Byte
   UserInvcG8V As Byte
   UserInvcG8F As Byte
   
   'Added just in case
   UserInvcG9 As Byte
   UserInvcG9E As Byte
   UserInvcG9V As Byte
   UserInvcG9F As Byte
   UserInvcG10 As Byte
   UserInvcG10E As Byte
   UserInvcG10V As Byte
   UserInvcG10F As Byte
   
   UserProd As Byte
   UserProdG1 As Byte
   UserProdG1E As Byte
   UserProdG1V As Byte
   UserProdG1F As Byte
   UserProdG2 As Byte
   UserProdG2E As Byte
   UserProdG2V As Byte
   UserProdG2F As Byte
   UserProdG3 As Byte
   UserProdG3E As Byte
   UserProdG3V As Byte
   UserProdG3F As Byte
   UserProdG4 As Byte
   UserProdG4E As Byte
   UserProdG4V As Byte
   UserProdG4F As Byte
   UserProdG5 As Byte
   UserProdG5E As Byte
   UserProdG5V As Byte
   UserProdG5F As Byte
   UserProdG6 As Byte
   UserProdG6E As Byte
   UserProdG6V As Byte
   UserProdG6F As Byte
   UserProdG7 As Byte
   UserProdG7E As Byte
   UserProdG7V As Byte
   UserProdG7F As Byte
   UserProdG8 As Byte
   UserProdG8E As Byte
   UserProdG8V As Byte
   UserProdG8F As Byte
   
   'Added just in case
   UserProdG9 As Byte
   UserProdG9E As Byte
   UserProdG9V As Byte
   UserProdG9F As Byte
   UserProdG10 As Byte
   UserProdG10E As Byte
   UserProdG10V As Byte
   UserProdG10F As Byte
   
   UserQual As Byte
   UserQualG1 As Byte
   UserQualG1E As Byte
   UserQualG1V As Byte
   UserQualG1F As Byte
   UserQualG2 As Byte
   UserQualG2E As Byte
   UserQualG2V As Byte
   UserQualG2F As Byte
   UserQualG3 As Byte
   UserQualG3E As Byte
   UserQualG3V As Byte
   UserQualG3F As Byte
   UserQualG4 As Byte
   UserQualG4E As Byte
   UserQualG4V As Byte
   UserQualG4F As Byte
   UserQualG5 As Byte
   UserQualG5E As Byte
   UserQualG5V As Byte
   UserQualG5F As Byte
   UserQualG6 As Byte
   UserQualG6E As Byte
   UserQualG6V As Byte
   UserQualG6F As Byte
   UserQualG7 As Byte
   UserQualG7E As Byte
   UserQualG7V As Byte
   UserQualG7F As Byte
   UserQualG8 As Byte
   UserQualG8E As Byte
   UserQualG8V As Byte
   UserQualG8F As Byte
   
   UserSale As Byte
   UserSaleG1 As Byte
   UserSaleG1E As Byte
   UserSaleG1V As Byte
   UserSaleG1F As Byte
   UserSaleG2 As Byte
   UserSaleG2E As Byte
   UserSaleG2V As Byte
   UserSaleG2F As Byte
   UserSaleG3 As Byte
   UserSaleG3E As Byte
   UserSaleG3V As Byte
   UserSaleG3F As Byte
   UserSaleG4 As Byte
   UserSaleG4E As Byte
   UserSaleG4V As Byte
   UserSaleG4F As Byte
   UserSaleG5 As Byte
   UserSaleG5E As Byte
   UserSaleG5V As Byte
   UserSaleG5F As Byte
   UserSaleG6 As Byte
   UserSaleG6E As Byte
   UserSaleG6V As Byte
   UserSaleG6F As Byte
   UserSaleG7 As Byte
   UserSaleG7E As Byte
   UserSaleG7V As Byte
   UserSaleG7F As Byte
   UserSaleG8 As Byte
   UserSaleG8E As Byte
   UserSaleG8V As Byte
   UserSaleG8F As Byte
   
   UserTime As Byte
   UserTimeG1 As Byte
   UserTimeG1E As Byte
   UserTimeG1V As Byte
   UserTimeG1F As Byte
   UserTimeG2 As Byte
   UserTimeG2E As Byte
   UserTimeG2V As Byte
   UserTimeG2F As Byte
   UserSparG3 As Byte
   UserSparG3E As Byte
   UserSparG3V As Byte
   UserSparG3F As Byte
   UserSparG4 As Byte
   UserSparG4E As Byte
   UserSparG4V As Byte
   UserSparG4F As Byte
   UserSparG5 As Byte
   UserSparG5E As Byte
   UserSparG5V As Byte
   UserSparG5F As Byte
   UserSparG6 As Byte
   UserSparG6E As Byte
   UserSparG6V As Byte
   UserSparG6F As Byte
   UserSparG7 As Byte
   UserSparG7E As Byte
   UserSparG7V As Byte
   UserSparG7F As Byte
   UserSparG8 As Byte
   UserSparG8E As Byte
   UserSparG8V As Byte
   UserSparG8F As Byte
   'UserZHideModule As Byte

End Type

Public Secure As SecurityClass


'Encrypts a new password
'Moderate security, but should do the job except for
'the most proficient.  They are dangerous anyway
'PassWord = ScramblePs(SomeNewOne)

Public Function ScramblePw(PassW As String) As String
   Dim a As Integer
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
      a = a + 1
      '1 extra byte
      Mid(NewPassWord, iList, 1) = Chr$(Asc(Mid$(PassW, a, 1)) + 1)
      If a = b Then Exit For
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
   Dim a As Integer
   Dim b As Integer
   Dim C As Integer
   Dim TempPw As String
   
   On Error GoTo ModErr1
   
   'Turn it around
   For a = 40 To 1 Step -1
      TempPw = TempPw & Mid$(PassWord, a, 1)
   Next
   C = Val(Right(TempPw, 2)) - 2
   For a = 3 To C Step 2
      GetSecPassword = GetSecPassword & Chr$(Asc(Mid$(TempPw, a, 1)) - 1)
   Next
   Exit Function
   
ModErr1:
   GetSecPassword = ""
   
End Function



Public Function CheckUserPassword(UserId As String, PwTest As String) As Byte
   Dim bTry As Byte
   Dim b As Boolean
   Dim iList As Integer
   Dim iFreeFile As Integer
   Dim sCoServer As String
   Dim sUserTest As String
   Dim sPwTest As String
   
   sUserTest = UCase$(UserId)
   sUserTest = GetSetting("ES2000", "Develop", "DevKey", "1225")
   '    If sUserTest = "1225" Then sCoServer = "c:\esi2000\" _
   '        Else sCoServer = App.Path & "\"
   sCoServer = App.Path & "\" '8/3/07 - always run in app dir
   iFreeFile = FreeFile
   On Error GoTo DiaErr1
   Open sCoServer & "rstval.eid" For Random Shared As iFreeFile Len = Len(SecPw)
   For iList = 1 To LOF(iFreeFile) \ Len(SecPw)
      Get #iFreeFile, , SecPw
      If Trim(SecPw.UserUcName) = sUserTest Then
         b = True
         Exit For
      End If
   Next
   If b Then
      sPwTest = GetSecPassword(SecPw.PassWord)
      If sPwTest = PwTest Then
         iList = SecPw.UserRecord
         SaveSetting "Esi2000", "System", "UserProfileRec", iList
         CheckUserPassword = 1
      Else
         CheckUserPassword = 0
      End If
   End If
   Exit Function
   
DiaErr1:
   bTry = bTry + 1
   If bTry < 3 Then Resume Next
   
End Function

'See If It's Active

Public Function CheckSecuritySettings(Optional FromMom As Boolean) As Byte
   Dim b As Byte
   Dim bTry As Byte
   Dim iLen As Integer
   Dim iFreeFile As Integer
   Dim iFreeFile2 As Integer
   Dim iUserRec As Integer
   Dim sCoServer As String
   MouseCursor 13
   If Trim(cUR.CurrentUser) = "" Then cUR.CurrentUser = GetSetting("Esi2000", "system", "UserId", cUR.CurrentUser)
   'If UCase$(cur.CurrentUser) = "LARRYH" Or UCase$(cur.CurrentUser) = "ESIADMN" Then
   If UCase$(cUR.CurrentUser) = "ESIADMN" Then
      bSecSet = 1
      '        User.Group1 = 1
      '        User.Group2 = 1
      '        User.Group3 = 1
      '        User.Group4 = 1
      '        User.Group5 = 1
      '        User.Group6 = 1
      '        User.Group7 = 1
      '        User.Group8 = 1
      
      InitializePermissions Secure, 1
      CheckSecuritySettings = 1
      Exit Function
   End If
   bSecSet = 0
   'Full Permissions for the Programmer
    bIdeRunning = IIf(RunningInIDE(), 1, 0)
   If bIdeRunning = 1 Then
      bSecSet = 1
      '        User.Group1 = 1
      '        User.Group2 = 1
      '        User.Group3 = 1
      '        User.Group4 = 1
      '        User.Group5 = 1
      '        User.Group6 = 1
      '        sCoServer = "c:\esi2000\"
      '    Else
      '        sCoServer = sFilePath
      InitializePermissions Secure, 1
   End If
   sCoServer = sFilePath '8/3/07 always run in app dir
   On Error GoTo ModErr1
   'b = 0 'testing
   If b <> 1 Then
      iFreeFile = FreeFile
      Open sCoServer & "rstval.edd" For Random Shared As iFreeFile Len = Len(Secure)
      iLen = LOF(iFreeFile) \ Len(Secure)
      If iLen > 0 Then
         iFreeFile2 = FreeFile
         Open sCoServer & "rstval.eid" For Random Shared As iFreeFile2 Len = Len(SecPw)
      Else
         If Err = 0 Then
            Close #iFreeFile
         End If
      End If
   Else
      iFreeFile = FreeFile
      Open sCoServer & "rstval.edd" For Random Shared As iFreeFile Len = Len(Secure)
      iFreeFile2 = FreeFile
      Open sCoServer & "rstval.eid" For Random Shared As iFreeFile2 Len = Len(SecPw)
      iLen = LOF(iFreeFile) \ Len(Secure)
   End If
   'If iLen > 3 Then       'always advanced security
   bSecSet = 1
   If FromMom Then
      CheckSecuritySettings = 1
      On Error Resume Next
      Close #iFreeFile
      Close #iFreeFile2
      MouseCursor 0
      Exit Function
   End If
   If b = 1 Then
      iUserRec = 2
   Else
      iUserRec = GetSetting("Esi2000", "System", "UserProfileRec", iUserRec)
   End If
   If iUserRec > 0 Then
      Get #iFreeFile, iUserRec, Secure
      Get #iFreeFile2, iUserRec, SecPw
      SetUserModulePermissions Secure
      CheckSecuritySettings = 1
   Else
      CheckSecuritySettings = 0
      '            If iLen > 3 Then       'redundant
      '                bSecSet = 1
      '            Else
      '                bSecSet = 0
      '            End If
   End If
   'Else
   '    CheckSecuritySettings = 0
   '    bSecSet = 0
   'End If
   On Error Resume Next
   Close #iFreeFile
   Close #iFreeFile2
   MouseCursor 0
   Exit Function
   
ModErr1:
   bTry = bTry + 1
   If bTry < 3 Then
      Resume Next
   Else
      CheckSecuritySettings = 0
      On Error GoTo 0
      MouseCursor 0
   End If
End Function

Public Function GetUser(UserId As String, PassWord As String) As Byte
   Dim bTry As Byte
   Dim iList As Integer
   Dim iFreeFile1 As Integer
   Dim iFreeFile2 As Integer
   Dim iLen As Integer
   Dim sNewPw As String
   If PassWord = "" Then
      GetUser = 0
      Exit Function
   End If
   
   On Error GoTo ModErr1
   '    If bIdeRunning = 1 Then
   '        sFilePath = "c:\esi2000\"
   '        'sFilePath = App.Path & "\"
   '    End If
   sFilePath = App.Path & "\" '8/3/07 - always run in local directory
   
   If Len(Trim(sFilePath)) <> "" Then
      iFreeFile1 = FreeFile
      Open sFilePath & "rstval.eid" For Random Shared As iFreeFile1 Len = Len(SecPw)
      iFreeFile2 = FreeFile
      Open sFilePath & "rstval.edd" For Random Shared As iFreeFile2 Len = Len(Secure)
      iLen = LOF(iFreeFile1) \ Len(SecPw)
      For iList = 2 To iLen
         Get #iFreeFile1, iList, SecPw
         If UCase$(UserId) = Trim(UCase$(SecPw.UserUcName)) Then
            'EsiLogon.txtUsr = Trim(SecPw.UserLcName)
            sNewPw = GetSecPassword(SecPw.PassWord)
            If sNewPw = PassWord Then
               Get #iFreeFile2, iList, Secure
               SetUserModulePermissions Secure
               sInitials = Secure.UserInitials
               sInitials = Trim(sInitials)
               SaveSetting "Esi2000", "System", "UserInitials", sInitials
               If Secure.UserActive = 1 Then
                  SaveSetting "Esi2000", "System", "UserProfileRec", iList
                  GetUser = 1
               Else
                  MsgBox "This User Is Marked As Inactive.", _
                     vbInformation, sSysCaption
                  SaveSetting "Esi2000", "System", "UserProfileRec", 0
                  GetUser = 0
               End If
            Else
               InitializePermissions Secure, 0
               SaveSetting "Esi2000", "System", "UserProfileRec", 0
               GetUser = 0
            End If
            Exit For
         End If
      Next
      Close #iFreeFile1
      Close #iFreeFile2
   Else
      MsgBox "This User Must Be Completely Set Up First.", _
         vbInformation, sSysCaption
      GetUser = 0
   End If
   Exit Function
   
ModErr1:
   bTry = bTry + 1
   If bTry < 3 Then Resume Next Else GetUser = 0
   
End Function

'9/5/06
'they didn't sign up, but we want them to see

Public Sub GroupTabPermissions(frm As Form)
   frm.lstSelect.Enabled = False
   frm.lblCustomer.ForeColor = ES_BLUE
   frm.lblCustomer.Visible = True
   frm.lblCustomer.ToolTipText = "This Group Is Not Part Of Your Company Feature Package"
   
End Sub

Public Sub InitializePermissions(obj As SecurityClass, DefaultValue As Byte)
   obj.UserAdmn = DefaultValue
   obj.UserAdmnG1 = DefaultValue
   obj.UserAdmnG1E = DefaultValue
   obj.UserAdmnG1V = DefaultValue
   obj.UserAdmnG1F = DefaultValue
   
   obj.UserAdmnG2 = DefaultValue
   obj.UserAdmnG2E = DefaultValue
   obj.UserAdmnG2V = DefaultValue
   obj.UserAdmnG2F = DefaultValue
   
   obj.UserAdmnG3 = DefaultValue
   obj.UserAdmnG3E = DefaultValue
   obj.UserAdmnG3V = DefaultValue
   obj.UserAdmnG3F = DefaultValue
   
   obj.UserAdmnG4 = DefaultValue
   obj.UserAdmnG4E = DefaultValue
   obj.UserAdmnG4V = DefaultValue
   obj.UserAdmnG4F = DefaultValue
   
   obj.UserAdmnG5 = DefaultValue
   obj.UserAdmnG5E = DefaultValue
   obj.UserAdmnG5V = DefaultValue
   obj.UserAdmnG5F = DefaultValue
   
   obj.UserAdmnG6 = DefaultValue
   obj.UserAdmnG6E = DefaultValue
   obj.UserAdmnG6V = DefaultValue
   obj.UserAdmnG6F = DefaultValue
   
   obj.UserAdmnG7 = DefaultValue
   obj.UserAdmnG7E = DefaultValue
   obj.UserAdmnG7V = DefaultValue
   obj.UserAdmnG7F = DefaultValue
   
   obj.UserAdmnG8 = DefaultValue
   obj.UserAdmnG8E = DefaultValue
   obj.UserAdmnG8V = DefaultValue
   obj.UserAdmnG8F = DefaultValue
   
   obj.UserEngr = DefaultValue
   obj.UserEngrG1 = DefaultValue
   obj.UserEngrG1E = DefaultValue
   obj.UserEngrG1V = DefaultValue
   obj.UserEngrG1F = DefaultValue
   
   obj.UserEngrG2 = DefaultValue
   obj.UserEngrG2E = DefaultValue
   obj.UserEngrG2V = DefaultValue
   obj.UserEngrG2F = DefaultValue
   
   obj.UserEngrG3 = DefaultValue
   obj.UserEngrG3E = DefaultValue
   obj.UserEngrG3V = DefaultValue
   obj.UserEngrG3F = DefaultValue
   
   obj.UserEngrG4 = DefaultValue
   obj.UserEngrG4E = DefaultValue
   obj.UserEngrG4V = DefaultValue
   obj.UserEngrG4F = DefaultValue
   
   obj.UserEngrG5 = DefaultValue
   obj.UserEngrG5E = DefaultValue
   obj.UserEngrG5V = DefaultValue
   obj.UserEngrG5F = DefaultValue
   
   obj.UserEngrG6 = DefaultValue
   obj.UserEngrG6E = DefaultValue
   obj.UserEngrG6V = DefaultValue
   obj.UserEngrG6F = DefaultValue
   
   obj.UserEngrG7 = DefaultValue
   obj.UserEngrG7E = DefaultValue
   obj.UserEngrG7V = DefaultValue
   obj.UserEngrG7F = DefaultValue
   
   obj.UserEngrG8 = DefaultValue
   obj.UserEngrG8E = DefaultValue
   obj.UserEngrG8V = DefaultValue
   obj.UserEngrG8F = DefaultValue
   
   obj.UserFina = DefaultValue
   obj.UserFinaG1 = DefaultValue
   obj.UserFinaG1E = DefaultValue
   obj.UserFinaG1V = DefaultValue
   obj.UserFinaG1F = DefaultValue
   
   obj.UserFinaG2 = DefaultValue
   obj.UserFinaG2E = DefaultValue
   obj.UserFinaG2V = DefaultValue
   obj.UserFinaG2F = DefaultValue
   
   obj.UserFinaG3 = DefaultValue
   obj.UserFinaG3E = DefaultValue
   obj.UserFinaG3V = DefaultValue
   obj.UserFinaG3F = DefaultValue
   
   obj.UserFinaG4 = DefaultValue
   obj.UserFinaG4E = DefaultValue
   obj.UserFinaG4V = DefaultValue
   obj.UserFinaG4F = DefaultValue
   
   obj.UserFinaG5 = DefaultValue
   obj.UserFinaG5E = DefaultValue
   obj.UserFinaG5V = DefaultValue
   obj.UserFinaG5F = DefaultValue
   
   obj.UserFinaG6 = DefaultValue
   obj.UserFinaG6E = DefaultValue
   obj.UserFinaG6V = DefaultValue
   obj.UserFinaG6F = DefaultValue
   
   obj.UserFinaG7 = DefaultValue
   obj.UserFinaG7E = DefaultValue
   obj.UserFinaG7V = DefaultValue
   obj.UserFinaG7F = DefaultValue
   
   obj.UserFinaG8 = DefaultValue
   obj.UserFinaG8E = DefaultValue
   obj.UserFinaG8V = DefaultValue
   obj.UserFinaG8F = DefaultValue
   
   obj.UserFinaG9 = DefaultValue
   obj.UserFinaG9E = DefaultValue
   obj.UserFinaG9V = DefaultValue
   obj.UserFinaG9F = DefaultValue
   
   obj.UserFinaG10 = DefaultValue
   obj.UserFinaG10E = DefaultValue
   obj.UserFinaG10V = DefaultValue
   obj.UserFinaG10F = DefaultValue
   
   obj.UserInvc = DefaultValue
   obj.UserInvcG1 = DefaultValue
   obj.UserInvcG1E = DefaultValue
   obj.UserInvcG1V = DefaultValue
   obj.UserInvcG1F = DefaultValue
   
   obj.UserInvcG2 = DefaultValue
   obj.UserInvcG2E = DefaultValue
   obj.UserInvcG2V = DefaultValue
   obj.UserInvcG2F = DefaultValue
   
   obj.UserInvcG3 = DefaultValue
   obj.UserInvcG3E = DefaultValue
   obj.UserInvcG3V = DefaultValue
   obj.UserInvcG3F = DefaultValue
   
   obj.UserInvcG4 = DefaultValue
   obj.UserInvcG4E = DefaultValue
   obj.UserInvcG4V = DefaultValue
   obj.UserInvcG4F = DefaultValue
   
   obj.UserInvcG5 = DefaultValue
   obj.UserInvcG5E = DefaultValue
   obj.UserInvcG5V = DefaultValue
   obj.UserInvcG5F = DefaultValue
   
   obj.UserInvcG6 = DefaultValue
   obj.UserInvcG6E = DefaultValue
   obj.UserInvcG6V = DefaultValue
   obj.UserInvcG6F = DefaultValue
   
   obj.UserInvcG7 = DefaultValue
   obj.UserInvcG7E = DefaultValue
   obj.UserInvcG7V = DefaultValue
   obj.UserInvcG7F = DefaultValue
   
   obj.UserInvcG8 = DefaultValue
   obj.UserInvcG8E = DefaultValue
   obj.UserInvcG8V = DefaultValue
   obj.UserInvcG8F = DefaultValue
   
   obj.UserInvcG9 = DefaultValue
   obj.UserInvcG9E = DefaultValue
   obj.UserInvcG9V = DefaultValue
   obj.UserInvcG9F = DefaultValue
   
   obj.UserInvcG10 = DefaultValue
   obj.UserInvcG10E = DefaultValue
   obj.UserInvcG10V = DefaultValue
   obj.UserInvcG10F = DefaultValue
   
   obj.UserProd = DefaultValue
   obj.UserProdG1 = DefaultValue
   obj.UserProdG1E = DefaultValue
   obj.UserProdG1V = DefaultValue
   obj.UserProdG1F = DefaultValue
   
   obj.UserProdG2 = DefaultValue
   obj.UserProdG2E = DefaultValue
   obj.UserProdG2V = DefaultValue
   obj.UserProdG2F = DefaultValue
   
   obj.UserProdG3 = DefaultValue
   obj.UserProdG3E = DefaultValue
   obj.UserProdG3V = DefaultValue
   obj.UserProdG3F = DefaultValue
   
   obj.UserProdG4 = DefaultValue
   obj.UserProdG4E = DefaultValue
   obj.UserProdG4V = DefaultValue
   obj.UserProdG4F = DefaultValue
   
   obj.UserProdG5 = DefaultValue
   obj.UserProdG5E = DefaultValue
   obj.UserProdG5V = DefaultValue
   obj.UserProdG5F = DefaultValue
   
   obj.UserProdG6 = DefaultValue
   obj.UserProdG6E = DefaultValue
   obj.UserProdG6V = DefaultValue
   obj.UserProdG6F = DefaultValue
   
   obj.UserProdG7 = DefaultValue
   obj.UserProdG7E = DefaultValue
   obj.UserProdG7V = DefaultValue
   obj.UserProdG7F = DefaultValue
   
   obj.UserProdG8 = DefaultValue
   obj.UserProdG8E = DefaultValue
   obj.UserProdG8V = DefaultValue
   obj.UserProdG8F = DefaultValue
   
   obj.UserProdG9 = DefaultValue
   obj.UserProdG9E = DefaultValue
   obj.UserProdG9V = DefaultValue
   obj.UserProdG9F = DefaultValue
   
   obj.UserProdG10 = DefaultValue
   obj.UserProdG10E = DefaultValue
   obj.UserProdG10V = DefaultValue
   obj.UserProdG10F = DefaultValue
   
   obj.UserQual = DefaultValue
   obj.UserQualG1 = DefaultValue
   obj.UserQualG1E = DefaultValue
   obj.UserQualG1V = DefaultValue
   obj.UserQualG1F = DefaultValue
   
   obj.UserQualG2 = DefaultValue
   obj.UserQualG2E = DefaultValue
   obj.UserQualG2V = DefaultValue
   obj.UserQualG2F = DefaultValue
   
   obj.UserQualG3 = DefaultValue
   obj.UserQualG3E = DefaultValue
   obj.UserQualG3V = DefaultValue
   obj.UserQualG3F = DefaultValue
   
   obj.UserQualG4 = DefaultValue
   obj.UserQualG4E = DefaultValue
   obj.UserQualG4V = DefaultValue
   obj.UserQualG4F = DefaultValue
   
   obj.UserQualG5 = DefaultValue
   obj.UserQualG5E = DefaultValue
   obj.UserQualG5V = DefaultValue
   obj.UserQualG5F = DefaultValue
   
   obj.UserQualG6 = DefaultValue
   obj.UserQualG6E = DefaultValue
   obj.UserQualG6V = DefaultValue
   obj.UserQualG6F = DefaultValue
   
   obj.UserQualG7 = DefaultValue
   obj.UserQualG7E = DefaultValue
   obj.UserQualG7V = DefaultValue
   obj.UserQualG7F = DefaultValue
   
   obj.UserQualG8 = DefaultValue
   obj.UserQualG8E = DefaultValue
   obj.UserQualG8V = DefaultValue
   obj.UserQualG8F = DefaultValue
   
   obj.UserSale = DefaultValue
   obj.UserSaleG1 = DefaultValue
   obj.UserSaleG1E = DefaultValue
   obj.UserSaleG1V = DefaultValue
   obj.UserSaleG1F = DefaultValue
   
   obj.UserSaleG2 = DefaultValue
   obj.UserSaleG2E = DefaultValue
   obj.UserSaleG2V = DefaultValue
   obj.UserSaleG2F = DefaultValue
   
   obj.UserSaleG3 = DefaultValue
   obj.UserSaleG3E = DefaultValue
   obj.UserSaleG3V = DefaultValue
   obj.UserSaleG3F = DefaultValue
   
   obj.UserSaleG4 = DefaultValue
   obj.UserSaleG4E = DefaultValue
   obj.UserSaleG4V = DefaultValue
   obj.UserSaleG4F = DefaultValue
   
   obj.UserSaleG5 = DefaultValue
   obj.UserSaleG5E = DefaultValue
   obj.UserSaleG5V = DefaultValue
   obj.UserSaleG5F = DefaultValue
   
   obj.UserSaleG6 = DefaultValue
   obj.UserSaleG6E = DefaultValue
   obj.UserSaleG6V = DefaultValue
   obj.UserSaleG6F = DefaultValue
   
   obj.UserSaleG7 = DefaultValue
   obj.UserSaleG7E = DefaultValue
   obj.UserSaleG7V = DefaultValue
   obj.UserSaleG7F = DefaultValue
   
   obj.UserSaleG8 = DefaultValue
   obj.UserSaleG8E = DefaultValue
   obj.UserSaleG8V = DefaultValue
   obj.UserSaleG8F = DefaultValue
   
   obj.UserTime = DefaultValue
   obj.UserTimeG1 = DefaultValue
   obj.UserTimeG1E = DefaultValue
   obj.UserTimeG1V = DefaultValue
   obj.UserTimeG1F = DefaultValue
   obj.UserTimeG2 = DefaultValue
   obj.UserTimeG2E = DefaultValue
   obj.UserTimeG2V = DefaultValue
   obj.UserTimeG2F = DefaultValue
   obj.UserSparG3 = DefaultValue
   obj.UserSparG3E = DefaultValue
   obj.UserSparG3V = DefaultValue
   obj.UserSparG3F = DefaultValue
   obj.UserSparG4 = DefaultValue
   obj.UserSparG4E = DefaultValue
   obj.UserSparG4V = DefaultValue
   obj.UserSparG4F = DefaultValue
   obj.UserSparG5 = DefaultValue
   obj.UserSparG5E = DefaultValue
   obj.UserSparG5V = DefaultValue
   obj.UserSparG5F = DefaultValue
   obj.UserSparG6 = DefaultValue
   obj.UserSparG6E = DefaultValue
   obj.UserSparG6V = DefaultValue
   obj.UserSparG6F = DefaultValue
   obj.UserSparG7 = DefaultValue
   obj.UserSparG7E = DefaultValue
   obj.UserSparG7V = DefaultValue
   obj.UserSparG7F = DefaultValue
   obj.UserSparG8 = DefaultValue
   obj.UserSparG8E = DefaultValue
   obj.UserSparG8V = DefaultValue
   obj.UserSparG8F = DefaultValue
   
   SetUserModulePermissions obj
   
End Sub

Public Sub SetUserModulePermissions(Secure As SecurityClass)
   bSections(SECTION_ADMN) = Secure.UserAdmn
   bSections(SECTION_PROD) = Secure.UserProd
   bSections(SECTION_ENGR) = Secure.UserEngr
   bSections(SECTION_SALE) = Secure.UserSale
   bSections(SECTION_FINA) = Secure.UserFina
   bSections(SECTION_QUAL) = Secure.UserQual
   If bSections(SECTION_QUAL) = 0 Then
      Debug.Print "whoops"
   End If
   bSections(SECTION_INVC) = Secure.UserInvc
   bSections(SECTION_TIME) = Secure.UserTime
   bSections(SECTION_POM) = 1
   bSections(SECTION_DATACOL) = 1
End Sub

Public Function GetHideModule() As Integer
   Dim RdoMod As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ISNULL(HideModuleButton, 0) HideModuleButton " _
          & " FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMod, ES_FORWARD)
   If bSqlRows Then
      With RdoMod
         GetHideModule = !HideModuleButton
         ClearResultSet RdoMod
      End With
   Else
        ' If the field is 0 then set to default 0
       GetHideModule = 0
   End If
   Set RdoMod = Nothing
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Public Function SetHideModule(ByVal iHideFlg As Integer)
    ' Update the Hide Module flag in preference table
    sSql = "UPDATE Preferences SET HideModuleButton ='" & CStr(iHideFlg) & "' WHERE PreRecord=1"
    clsADOCon.ExecuteSQL sSql ', rdExecDirect
    
End Function



Public Sub CopyUsersToTable()
   'returns table name
   Dim iFreeFile As Integer, iFreeFile2 As Integer
   Dim sCoServer As String
   clsADOCon.ExecuteSQL "delete from EsReportUserPermissions"
   clsADOCon.ExecuteSQL "delete from EsReportUsers"
   
   Dim u As SecurityArray, pw As PassWordClass
   Dim I As Integer, module As Integer, Section As Integer, perm As Integer
   iFreeFile = FreeFile
   'sCoServer = App.Path & "\"
   Open sFilePath & "rstval.edd" For Random Shared As iFreeFile Len = Len(Secure)
   iFreeFile2 = FreeFile
   Open sFilePath & "rstval.eid" For Random Shared As iFreeFile2 Len = Len(SecPw)
   For I = 1 To LOF(iFreeFile) \ Len(Secure)
      Get #iFreeFile, I, u
      Get #iFreeFile2, I, pw
      
If Z2B(pw.UserLcName) <> "" Then

      sSql = "insert into EsReportUsers" & vbCrLf _
         & "(UserID, UserName, Initials, NickName, Active, Created, Level, Admin)" & vbCrLf _
         & "values('" & Replace(Z2B(pw.UserLcName), "'", "''") & "'," _
         & "'" & Replace(Z2B(u.UserName), "'", "''") & "'," _
         & "'" & Replace(Z2B(u.UserInitials), "'", "''") & "'," _
         & "'" & Replace(Z2B(u.UserNickName), "'", "''") & "', " _
         & u.UserActive & ", " _
         & IIf(Z2B(u.UserNew) = "", "null", "'" & Z2B(u.UserNew) & "'") & ", " _
         & u.UserLevel & ", " _
         & pw.UserAdmn _
         & ")"
      clsADOCon.ExecuteSQL sSql

      'insert permissions for this user
      Dim modName As String, secName As String
      perm = 1    'subscript for permission array
      For module = 1 To ModuleCount
         modName = ModuleName(module)
'Debug.Print modName & ": " & perm
         If modName <> "UNKNOWN" Then
            For Section = 1 To SectionCount(module)
            secName = SectionName(module, Section)
            If secName <> "UNKNOWN" Then
               sSql = "insert into EsReportUserPermissions" & vbCrLf _
                  & "(UserID, UserName, ModuleName, SectionName, GroupPermission, EditPermission, ViewPermission, FunctionPermission)" & vbCrLf _
                  & "values('" & Replace(Z2B(pw.UserLcName), "'", "''") & "', '" & Replace(Z2B(u.UserName), "'", "''") & "', '" & Replace(modName, "'", "''") & "', '" & Replace(secName, "'", "''") & "', " _
                  & u.UserPermissions(perm) & ", " & u.UserPermissions(perm + 1) & ", " _
                  & u.UserPermissions(perm + 2) & ", " & u.UserPermissions(perm + 3) & ")"
                  clsADOCon.ExecuteSQL sSql
            End If
            perm = perm + 4
            Next Section
         End If
      Next module
End If
   Next I
   Close #iFreeFile
   
   '6/10/2009 Not need
   'MsgBox "Tables EsReportUsers and EsReportUserPermissions have been populated." & vbCrLf _
   '   & "Reports have not yet been written for this user data, but you can view it."
End Sub

Private Function Z2B(inputstring As String) As String
   'replace zeros with blanks
   Z2B = Trim(Replace(inputstring, Chr(0), " "))
End Function

Public Function ModuleName(ModuleNumber As Integer) As String
   'returns name of module 1 through 8
   'return = "UNKNOWN" if no such module
   Select Case ModuleNumber
   Case 1:
      ModuleName = "Administration"
   Case 2:
      ModuleName = "Engineering"
   Case 3:
      ModuleName = "Finance"
   Case 4:
      ModuleName = "Inventory Control"
   Case 5:
      ModuleName = "Production Control"
   Case 6:
      ModuleName = "Quality Assurance"
   Case 7:
      ModuleName = "Sales"
   Case 8:
      ModuleName = "Time Management"
   Case 9:
      ModuleName = "POM Module"
   Case Else
      ModuleName = "UNKNOWN"
   End Select

End Function

Public Function ModuleCount() As Integer
   ModuleCount = 8
End Function

Public Function SectionCount(ModuleNumber As Integer) As Integer
   Select Case ModuleNumber
   Case 1:
      SectionCount = 8     'admin
   Case 2:
      SectionCount = 8     'engr
   Case 3:
      SectionCount = 10    'finance
   Case 4:
      SectionCount = 10    'inv
   Case 5:
      SectionCount = 10    'prod
   Case 6:
      SectionCount = 8     'qual
   Case 7:
      SectionCount = 8     'sale
   Case 8:
      SectionCount = 2     'time
   Case Else
      SectionCount = 0
   End Select
End Function

Public Function SectionName(ModuleNumber As Integer, SectionNumber As Integer) As String
   'returns name of Section for a given module
   'return = "UNKNOWN" if no such Section
   
   'administration module
   Select Case ModuleNumber
   Case 1:
      Select Case SectionNumber
      Case 1:
         SectionName = "System"
      Case 2:
         SectionName = "Sales"
      Case 3:
         SectionName = "Production Control"
      Case 4:
         SectionName = "Time Management"
      Case 5:
         SectionName = "Inventory Management"
      Case 6:
         SectionName = "System Help"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'engineering module
   Case 2:
      Select Case SectionNumber
      Case 1:
         SectionName = "Routings"
      Case 2:
         SectionName = "Bills"
      Case 3:
         SectionName = "Documents"
      Case 4:
         SectionName = "Tooling"
      Case 5:
         SectionName = "Estimating"
      Case 6:
         SectionName = "UNKNOWN"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
      
   'finance module
   Case 3:
      Select Case SectionNumber
      Case 1:
         SectionName = "Accounts Receivable"
      Case 2:
         SectionName = "Accounts Payable"
      Case 3:
         SectionName = "General Ledger"
      Case 4:
         SectionName = "Journals"
      Case 5:
         SectionName = "Closing"
      Case 6:
         SectionName = "Product Costing"
      Case 7:
         SectionName = "Job Costing"
      Case 8:
         SectionName = "Sales Analysis"
      Case 9:
         SectionName = "UNKNOWN"
      Case 10:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'inventory control module
   Case 4:
      Select Case SectionNumber
      Case 1:
         SectionName = "Inventory"
      Case 2:
         SectionName = "Material"
      Case 3:
         SectionName = "Receiving"
      Case 4:
         SectionName = "Inventory Management"
      Case 5:
         SectionName = "Lot Tracking"
      Case 6:
         SectionName = "UNKNOWN"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case 9:
         SectionName = "UNKNOWN"
      Case 10:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'production control module
   Case 5:
      Select Case SectionNumber
      Case 1:
         SectionName = "MO's"
      Case 2:
         SectionName = "CRP"
      Case 3:
         SectionName = "Purchasing"
      Case 4:
         SectionName = "PAC"
      Case 5:
         SectionName = "MRP"
      Case 6:
         SectionName = "UNKNOWN"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case 9:
         SectionName = "UNKNOWN"
      Case 10:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'quality control module
   Case 6:
      Select Case SectionNumber
      Case 1:
         SectionName = "Inspection Reports"
      Case 2:
         SectionName = "First Article Inspection"
      Case 3:
         SectionName = "Statistical Process Control"
      Case 4:
         SectionName = "On Dock Inspection"
      Case 5:
         SectionName = "Administration"
      Case 6:
         SectionName = "UNKNOWN"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'Sales module
   Case 7:
      Select Case SectionNumber
      Case 1:
         SectionName = "Order Processing"
      Case 2:
         SectionName = "Packing Slips"
      Case 3:
         SectionName = "Bookings/Backlog"
      Case 4:
         SectionName = "Commissions"
      Case 5:
         SectionName = "UNKNOWN"
      Case 6:
         SectionName = "UNKNOWN"
      Case 7:
         SectionName = "UNKNOWN"
      Case 8:
         SectionName = "UNKNOWN"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
   
   'time management module
   Case 8:
      Select Case SectionNumber
      Case 1:
         SectionName = "Time Charges"
      Case 2:
         SectionName = "Time & Attendance"
      Case Else:
         SectionName = "UNKNOWN"
      End Select
      
   Case Else
      SectionName = "UNKNOWN"
   End Select

End Function

Public Function RunningInIDE() As Boolean
   'Test to see if we are in VB or user mode
   'Check to see where the program is running.
   'Assume that we are not running in VB5 for now
   'Calling Debug is ignored except in VB, so will produce
   'an error...
   RunningInIDE = False
   
   'Exit Function '@@@temp test
   
   On Error GoTo ModErr
   Debug.Print 1 / 0
   Exit Function
   
ModErr:
   On Error GoTo 0
   'yep, it's VB
   RunningInIDE = True
   'sFilePath = Mid(App.Path, 1, LastIndexOf(App.Path, "\")) + "ReleaseCandidates\"
   'App.Title = "ESI VB Run " & Left(MdiSect.Caption, 4)
   
End Function

Public Function InTestMode() As Boolean
   If Dir(App.Path & "\UseTestDatabase.txt") <> "" Then
      InTestMode = True
   End If
End Function
