Attribute VB_Name = "Registry"
Option Explicit

Public Enum eUserSettingType
   USERSETTING_UserLogin = 0
   USERSETTING_UserPassword = 1
   USERSETTING_SaveUserLogin = 2
   USERSETTING_SaveDatabase = 3
   USERSETTING_DefaultDatabase = 4
   USERSETTING_ServerName = 5
   USERSETTING_DatabaseName = 6
   USERSETTING_SqlPassword = 7
   USERSETTING_SqlLogin = 8
   USERSETTING_SqlDsn = 9
   USERSETTING_BackgroundColorRGB = 10
End Enum

Private OldApplications As Variant
Private OldKeys As Variant
Private NewKeys As Variant
Private DefaultValues As Variant
Private ConfigKeys As Variant

Private Const SectionName = "FUSION_USERSETTINGS"
Private Const APPKEYNAME = "Esi2000"
Private Const CURRENTSECTION = "Login\Current"
Private Const PATHSECTION = "Login"

Public Function GetConfigDBPass() As String
   Dim sPassword As String
   Dim sEncrypted As String
   
   GetConfigDBPass = ""
   sEncrypted = GetConfUserSetting(USERSETTING_SqlPassword)
   sPassword = GetSecPassword(sEncrypted)

   GetConfigDBPass = sPassword
   
End Function

Public Function GetDatabasePassword() As String
   Dim sPassword As String, sPassword2 As String
   
   'pre-version 8 - unencrypted password was stored in RegOne
   GetDatabasePassword = ""
   sPassword = GetSetting("SysCan", "System", "RegOne", "unknown_pwd")
   
   'post version 8
   If sPassword = "unknown_pwd" Then
      sPassword = GetUserSetting(USERSETTING_SqlPassword)
      If sPassword = "unknown_password" Then
         Exit Function
      Else
         sPassword = GetSecPassword(sPassword)
      End If
   
   'pre version 8
   Else
      If Len(sPassword) = 0 Then
         'blank password
      ElseIf Len(sPassword) <= 5 Then
         MsgBox "Esi2000.Registry.GetDatabasePassword failed to encrypt pre-version 8 password.  Contact ESI"
         Exit Function
      Else
         sPassword = Mid(sPassword, 4, Len(sPassword) - 5)
      End If
      PutDatabasePassword sPassword
      DeleteSetting "SysCan", "System", "RegOne"
   End If
   GetDatabasePassword = sPassword
End Function

Public Sub PutDatabasePassword(sPassword As String)
   'encrypt database password and store it in the registry
   
   Dim sEncrypted As String, sEncrypted2 As String, sPassword2 As String
   sEncrypted = ScramblePw(sPassword)
   'SaveSetting "SysCan", "System", "RegTwo", sEncrypted
   SaveUserSetting USERSETTING_SqlPassword, sEncrypted
   'sPassword2 = GetUserSetting(USERSETTING_SqlPassword)
   sEncrypted2 = GetUserSetting(USERSETTING_SqlPassword)
   sPassword2 = GetSecPassword(sEncrypted)
   If sPassword2 <> sPassword Then
      MsgBox "Esi2000.Registry.PutDatabasePassword encryption failed.  Contact ESI."
   End If
   
End Sub

Private Sub InitializeConfigArrays()
   On Error Resume Next
   ConfigKeys = Array("", "", "", "SAVEDATABASE", _
      "DEFAULTDATABASE", "SERVERNAME", "DATABASENAME", "SQLPASSWORD", "SQLLOGIN", "SQLDSN", "BACKGROUNDCOLORRGB")

End Sub

Private Sub InitializeRegistryArrays()
   On Error Resume Next
   If UBound(OldKeys) > 0 Then
      If Err = 0 Then
         Exit Sub
      End If
   End If
   OldApplications = Array("Esi2000", "Esi2000", "Esi2000", "Esi2000", _
      "Esi2000", "Esi2000", "Esi2000", "SysCan", "UserObjects", "Esi2000")
   OldKeys = Array("Lastuser", "LastId", "SaveId", "SaveUserDatabase", _
      "DefaultUserDatabase", "ServerID", "CurDatabase", "RegTwo", "NoReg", "SqlDsn")
   NewKeys = Array("UserLogin", "UserPassword", "SaveUserLogin", "SaveDatabase", _
      "DefaultDatabaseName", "ServerName", "DatabaseName", "SqlPassword", "SqlLogin", "SqlDsn")
   DefaultValues = Array("", "", "0", "0", _
      "", "", "", "", "", "")
      
End Sub

Public Function GetConfUserSetting(Setting As eUserSettingType) As String


   Dim X As Long
   Dim sSection As String, sEntry As String, sDefault As String
   Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
   Dim sValue As String
   Dim sIniPath As String

   On Error GoTo modErr1
   
   'search for setting in current path
   InitializeConfigArrays
   If RunningInIDE Then
      Dim sPath As String
      sPath = Mid$(App.Path, 1, InStrRev(App.Path, "\"))
      sPath = sPath & "ReleaseCandidates\FusionInit.ini"
      'MsgBox sPath
      sIniPath = sPath
   Else
      sIniPath = App.Path & "\" & "FusionInit.ini"
   End If
   
   
   If Dir(sIniPath) = "" Then
      MsgBox "Error: GetConfUserSetting: Couldn't find config file."
      End
   End If
   
   sSection = SectionName
   sEntry = ConfigKeys(Setting)
   sDefault = ""
   sRetBuf = String(256, vbNull) '256 null characters
   iLenBuf = Len(sRetBuf)
   sFileName = sIniPath
   X = GetPrivateProfileString(sSection, sEntry, _
                     "", sRetBuf, iLenBuf, sFileName)
   sValue = Trim(Left$(sRetBuf, X))
   
   If sValue <> "" Then
      GetConfUserSetting = sValue
   Else
      GetConfUserSetting = ""
   End If
   
   Exit Function
   
modErr1:
   MsgBox "Error: GetConfUserSetting :" & Err.Description

End Function


Public Function GetUserSetting(Setting As eUserSettingType) As String
   'get setting from registry
   
   'search for setting in current path
   InitializeRegistryArrays
   Dim pathKey As String, Value As String
   pathKey = PATHSECTION & "\" & App.Path
   
   'if running in vb, use general setting
   If RunningInIDE Then
      Value = ""
   Else
      Value = GetSetting(APPKEYNAME, pathKey, NewKeys(Setting), "")
   End If
   
   'if no setting for current path, get current general setting
   If Value = "" Then
      Value = GetSetting(APPKEYNAME, CURRENTSECTION, NewKeys(Setting), "")
      'value = ""     '@@@ TEST
      
      'if still no value, get value from old pre-R66 (10/16/08) registry entry
      If Value = "" Then
         Value = GetSetting(OldApplications(Setting), "System", OldKeys(Setting), DefaultValues(Setting))
         'value = DefaultValues(Setting)   '@@@ TEST
         
         'place the value the current setting
Debug.Print NewKeys(Setting) & " = " & Value
         SaveSetting APPKEYNAME, CURRENTSECTION, NewKeys(Setting), Value
      End If
      
      'place the value in the new path setting
      SaveSetting APPKEYNAME, pathKey, NewKeys(Setting), Value
   End If
   
   GetUserSetting = Value
End Function

Public Sub SaveUserSetting(Setting As eUserSettingType, Value As String)
   'save setting in registry
   
   'save setting in current path
   InitializeRegistryArrays
   Dim pathKey As String
   pathKey = PATHSECTION & "\" & App.Path
   SaveSetting APPKEYNAME, pathKey, NewKeys(Setting), Value
   
   'save setting in current section
   SaveSetting APPKEYNAME, CURRENTSECTION, NewKeys(Setting), Value
End Sub
