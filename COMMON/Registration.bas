Attribute VB_Name = "Registration"
Option Explicit

Private Type LicenseFileType
    EncProductKey As String * 30
    EncRegId As String * 6
    EncExpDate As Date
End Type

Const sLicenseFile As String = "es2000.lic"

Private lRegId As Long
Private sCompanyName As String
Private sProductKey As String
Private dExpDate As Date
Private iNumLicenses As Integer
Private iPOMLicenses As Integer
Private sAppDir As String

Dim LicenseFile As LicenseFileType

Private Function WriteOutLicenseFile() As Boolean
    Dim iFileNo As Integer
    Err = 0
    On Error Resume Next
    WriteOutLicenseFile = False
    
    sAppDir = App.Path
    If Not Right(sAppDir, 1) = "\" Then sAppDir = sAppDir & "\"
    iFileNo = FreeFile
    Open sAppDir & sLicenseFile For Binary Lock Read Write As #iFileNo
    LicenseFile.EncProductKey = Encrypt(sProductKey)
    LicenseFile.EncRegId = Encrypt(Trim(Str(lRegId)))
    LicenseFile.EncExpDate = dExpDate
                
    Put #iFileNo, , LicenseFile
    Close iFileNo
    If Err = 0 Then WriteOutLicenseFile = True
    
End Function

Private Function GetRegistrationInfo()
   Dim RdoRslt As ADODB.Recordset
   Dim iFileNo As Integer
   
   sAppDir = App.Path
   If Not Right(sAppDir, 1) = "\" Then sAppDir = sAppDir & "\"
   
   sSql = "SELECT REGID, PRODUCTKEY, CONAME FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRslt, ES_FORWARD)
   If bSqlRows Then
       lRegId = Val("" & RdoRslt!RegID)
       sCompanyName = Trim("" & RdoRslt!CONAME)
       sProductKey = Trim("" & RdoRslt!ProductKey)
   End If
   Set RdoRslt = Nothing
   
   If lRegId = 0 And Len(sProductKey) = 0 And Dir(sAppDir & sLicenseFile, vbHidden) = "" Then
       ' If not product key or registration id is installed and they haven't done this before,
       ' create a single user license that expires in 30 days and save it. This is from the initial
       ' creating of this logic. After this, we will not allow them to just wipe it out
       lRegId = Val(GenerateRandomString(6))
       
       dExpDate = Now + 30
       iNumLicenses = 1
       iPOMLicenses = 1
       sProductKey = CreateProductKey
       clsADOCon.ExecuteSQL "UPDATE ComnTable SET REGID= " & lRegId & " ,PRODUCTKEY = '" & Trim(sProductKey) & "' "
               
       iFileNo = FreeFile
       Open sAppDir & sLicenseFile For Binary Lock Read Write As #iFileNo
       LicenseFile.EncProductKey = Encrypt(sProductKey)
       LicenseFile.EncRegId = Encrypt(Trim(Str(lRegId)))
       LicenseFile.EncExpDate = dExpDate
               
       Put #iFileNo, , LicenseFile
       Close iFileNo
       If Not Dir(sAppDir & sLicenseFile) = "" Then SetAttr sAppDir & sLicenseFile, vbHidden
   End If

   iNumLicenses = LicensedFusionUsersAllowed
   iPOMLicenses = LicensedPOMUsersAllowed
   dExpDate = GetExpirationDate

End Function


Public Function GetRegCompanyName() As String
    If Len(sCompanyName) = 0 Then GetRegistrationInfo
    GetRegCompanyName = sCompanyName
End Function


Public Function RegDaysLeft() As Long
    If Not IsDate(dExpDate) Then GetRegistrationInfo
    RegDaysLeft = DateDiff("d", Now, dExpDate)
End Function

Public Function IsRegistrationExpired() As Boolean
    If RegDaysLeft < 1 Then IsRegistrationExpired = True Else IsRegistrationExpired = False
End Function

Public Function RegistrationID() As Long
    If lRegId = 0 Then GetRegistrationInfo
    RegistrationID = lRegId
End Function

Public Function ProductKey() As String
    If Len(sProductKey) = 0 Then GetRegistrationInfo
    ProductKey = sProductKey
End Function


Private Function GenerateRandomString(ByRef Length As Integer) As String
    Randomize
    Dim allowableChars As String
    allowableChars = "0123456789"

    Dim I As Integer
    For I = 1 To Length
        GenerateRandomString = GenerateRandomString & Mid$(allowableChars, Int(Rnd() * Len(allowableChars) + 1), 1)
    Next
End Function


Private Function Encrypt(ByVal icText As String) As String
Dim icLen As Integer
Dim icNewText As String
Dim icchar As String
Dim I As Integer


icchar = ""
   icLen = Len(icText)
   For I = 1 To icLen
       icchar = Mid(icText, I, 1)
       Select Case Asc(icchar)
           Case 65 To 90
               icchar = Chr(Asc(icchar) + 127)
           Case 97 To 122
               icchar = Chr(Asc(icchar) + 121)
           Case 48 To 57
               icchar = Chr(Asc(icchar) + 196)
           Case 32
               icchar = Chr(32)
       End Select
       icNewText = icNewText + icchar
   Next
   Encrypt = icNewText
End Function

Private Function Decrypt(ByVal icText As String) As String
Dim icLen As Integer
Dim icNewText As String
Dim icchar As String
Dim I As Integer

icchar = ""
   icLen = Len(icText)
   For I = 1 To icLen
       icchar = Mid(icText, I, 1)
       Select Case Asc(icchar)
           Case 192 To 217
               icchar = Chr(Asc(icchar) - 127)
           Case 218 To 243
               icchar = Chr(Asc(icchar) - 121)
           Case 244 To 253
               icchar = Chr(Asc(icchar) - 196)
           Case 32
               icchar = Chr(32)
       End Select
       icNewText = icNewText + icchar
   Next
   Decrypt = icNewText
End Function


Private Function ReverseString(ByVal sInComing As String) As String
    Dim I As Integer
    Dim sNewVal As String

    sNewVal = ""
    For I = Len(sInComing) To 1 Step -1
        sNewVal = sNewVal + Mid(sInComing, I, 1)
    Next I
    ReverseString = sNewVal
End Function



Private Function Dec2Hex(ByVal DecimalIn As Long) As String
  Dim X As Integer
  Dim BinaryString As String
  Const BinValues = "*0000*0001*0010*0011" & _
                    "*0100*0101*0110*0111" & _
                    "*1000*1001*1010*1011" & _
                    "*1100*1101*1110*1111*"
  Const HexValues = "0123456789ABCDEF"
  Const MaxNumOfBits As Long = 96
  BinaryString = ""
  Do While DecimalIn <> 0
    BinaryString = Trim$(Str$(DecimalIn - 2 * _
                   Int(DecimalIn / 2))) & BinaryString
    DecimalIn = Int(DecimalIn / 2)
  Loop
  BinaryString = String$((4 - Len(BinaryString) _
                 Mod 4) Mod 4, "0") & BinaryString
  For X = 1 To Len(BinaryString) - 3 Step 4
    Dec2Hex = Dec2Hex & Mid$(HexValues, _
              (4 + InStr(BinValues, "*" & _
              Mid$(BinaryString, X, 4) & "*")) \ 5, 1)
  Next
End Function

Private Function Hex2Dec(ByVal HexString As String) As Long
  Dim X As Integer
  Dim bInstr As String
  Const TwoToThe49thPower As String = "562949953421312"
  If Left$(HexString, 2) Like "&[hH]" Then
    HexString = Mid$(HexString, 3)
  End If
  If Len(HexString) <= 23 Then
    Const BinValues = "0000000100100011" & _
                      "0100010101100111" & _
                      "1000100110101011" & _
                      "1100110111101111"
    For X = 1 To Len(HexString)
      bInstr = bInstr & Mid$(BinValues, _
               4 * Val("&h" & Mid$(HexString, X, 1)) + 1, 4)
    Next
    Hex2Dec = CDec(0)
    For X = 0 To Len(bInstr) - 1
      If X < 50 Then
        Hex2Dec = Hex2Dec + Val(Mid(bInstr, _
                  Len(bInstr) - X, 1)) * 2 ^ X
      Else
        Hex2Dec = Hex2Dec + CDec(TwoToThe49thPower) * _
                  Val(Mid(bInstr, Len(bInstr) - X, 1)) * 2 ^ (X - 49)
      End If
    Next
  Else
    ' Number is too big, handle error here
  End If
End Function


Private Function JulianToDate(lJulianDate As Long) As String
    Dim lYear, lDay, lDaysInYear As Long
    If lJulianDate = 0 Then Exit Function
    On Error Resume Next
    
    lYear = lJulianDate \ 1000             'get the year part
    lDay = lJulianDate - lYear * 1000     'get the day of the year part
    lDaysInYear = DaysInTheYear(DateSerial(lYear, 1, 1)) 'number of days in the year
    If lDay >= 1 And lDay <= lDaysInYear Then         'within the range?
        JulianToDate = Format(DateSerial(lYear, 1, 1) + lDay - 1, "mm/dd/yyyy")     'yes, return what we found
    End If
    If Err > 0 Then JulianToDate = ""
End Function

Private Function DateToJulian(sDate As String) As Long
    Dim lYear, lDay, lJulianDate As Long
    
    lYear = Year(sDate)                'get the year part
    lDay = DateDiff("y", DateSerial(lYear, 1, 1), sDate) + 1  'day part
    DateToJulian = CLng(Right$(Format$(lYear, "0000"), 2) & Format$(lDay, "000"))   'convert to yyddd
End Function

Private Function DaysInTheYear(dDate As Date) As Long
    DaysInTheYear = DateDiff("y", DateSerial(Year(dDate), 1, 1), LastDayOfTheYear(dDate)) + 1
End Function


Private Function LastDayOfTheYear(dDate As Date) As Date
    LastDayOfTheYear = DateSerial(Year(dDate) + 1, 1, 1) - 1 'one day less than January 1 the next year
End Function

Private Function CreateLastPartOfKey(ByVal sProdKeySoFar As String)
    Dim lTemp As Long
    Dim I As Integer
    Dim sTemp As String
    sTemp = ""
    
    lTemp = 0
    For I = 1 To 5
        lTemp = lTemp + Asc(Mid(sProdKeySoFar, I, 1))
    Next I
    sTemp = sTemp & Right("00" & Dec2Hex(lTemp), 2)
    
    lTemp = 0
    For I = 7 To 11
        lTemp = lTemp + Asc(Mid(sProdKeySoFar, I, 1))
    Next I
    sTemp = sTemp & Right("00" & Dec2Hex(lTemp), 2)
    
    lTemp = 0
    For I = 13 To 17
        lTemp = lTemp + Asc(Mid(sProdKeySoFar, I, 1))
    Next I
    sTemp = sTemp & Right("00" & Dec2Hex(lTemp), 1)
     
    CreateLastPartOfKey = sTemp
End Function

Private Function IsInvalidProductKey() As Boolean
    Dim sTemp As String
    If Len(sProductKey) = 0 Then
        GetRegistrationInfo
        If Len(sProductKey) = 0 Then
            IsInvalidProductKey = True
            Exit Function
        End If
    End If
    sTemp = CreateProductKey
    If Left(sTemp, Len(sTemp)) = Left(sProductKey, Len(sTemp)) Then IsInvalidProductKey = False Else IsInvalidProductKey = True
End Function


Public Function LicensedFusionUsersAllowed(Optional sProdKey As String) As Integer
    Dim lTemp, lTemp2 As Long
    Dim sTemp As String
    
    If Len(sProductKey) = 0 Or (lRegId = 0) Then
        GetRegistrationInfo
        If Len(sProductKey) = 0 Then
            LicensedFusionUsersAllowed = 0
            Exit Function
        End If
    End If
    
    lTemp2 = 65
    If Len(sProductKey) > 5 Then sTemp = Mid(sProductKey, 13, 5) Else sTemp = sProductKey
    If Len(sProdKey) > 0 Then sTemp = Mid(sProdKey, 13, 5)
    lTemp = Hex2Dec(ReverseString(sTemp))
    LicensedFusionUsersAllowed = (lTemp - lRegId) / lTemp2
End Function

Public Function LicensedPOMUsersAllowed(Optional sProdKey As String) As Integer
    Dim sTemp As String
    Dim iTemp As Integer
    
    
    If Len(sProductKey) = 0 Or lRegId = 0 Then GetRegistrationInfo
    
    If Len(sProdKey) = 0 Then sTemp = Mid(sProductKey, 3, 3) Else sTemp = Mid(sProdKey, 3, 3)
    iTemp = Val(Right(LTrim(Str(lRegId)), 2))
    
    LicensedPOMUsersAllowed = (Val(Hex2Dec(sTemp)) - iTemp)
End Function


Public Function GetExpirationDate(Optional sProdKey As String) As Date
    Dim sTemp As String
    Dim lJulianDate As Long
    On Error Resume Next
    
    If Len(sProductKey) = 0 And Len(sProdKey) = 0 Then
        GetRegistrationInfo
        If Len(sProductKey) = 0 Then
            GetExpirationDate = Now - 1
            Exit Function
        End If
    End If
    
    If Len(sProductKey) > 5 Then sTemp = Mid(sProductKey, 7, 5) Else sTemp = sProductKey
    If Len(sProdKey) > 0 Then sTemp = Mid(sProdKey, 7, 5)
    sTemp = ReverseString(sTemp)
    GetExpirationDate = JulianToDate(Val(sTemp))
End Function


Private Function CreateFirstPartOfKey() As String
    Dim lFirstPart As Long
    Dim sTemp As String
    Dim lTemp As Long
        
    If Len(sCompanyName) = 0 Or lRegId = 0 Or iPOMLicenses = 0 Then GetRegistrationInfo
    
    lFirstPart = (Asc(Left(sCompanyName, 1)) + Asc(Right(sCompanyName, 1)))
    sTemp = Right(LTrim(Str(lRegId)), 2)
    lTemp = Val(sTemp) + iPOMLicenses
    sTemp = Right("000" & Dec2Hex(lTemp), 3)
    
    CreateFirstPartOfKey = Right("00" & Dec2Hex(lFirstPart), 2) + sTemp
End Function


Public Function CreateProductKey(Optional sCoName As String, Optional lRegistrationID As Long, Optional iNoLic As Integer, Optional iPOMLic As Integer, Optional dExpirationDte As Date) As String
    Dim lFirstPart As Long
    Dim sTemp As String
    Dim sProdKey As String
    If Len(sCoName) > 0 And lRegistrationID > 0 Then
        lRegId = lRegistrationID
        sCompanyName = sCoName
        dExpDate = dExpirationDte
        iNumLicenses = iNoLic
        iPOMLicenses = iPOMLic
    Else
        If Len(sCompanyName) = 0 Then GetRegistrationInfo
    End If

    sProdKey = CreateFirstPartOfKey & "-" & CreateDatePartOfKey & "-" & CreateLicKeyPartOfKey & "-"
    CreateProductKey = sProdKey & CreateLastPartOfKey(sProdKey)
End Function



Public Function ValidProductKey() As Boolean
    Dim iFileNo As Integer
    Dim sTemp As String
    Dim dTemp As Date
    Dim lTemp As Long
    
    If Len(sProductKey) = 0 Then GetRegistrationInfo
    If Len(Trim(sProductKey)) = 0 Or IsInvalidProductKey Then
        ValidProductKey = False
        Exit Function
    End If
    
    iFileNo = FreeFile
    Open sAppDir & sLicenseFile For Binary Lock Read Write As #iFileNo
    Get #iFileNo, , LicenseFile
    sTemp = Decrypt(LicenseFile.EncProductKey)
    lTemp = Val(Decrypt(LicenseFile.EncRegId))
    dTemp = LicenseFile.EncExpDate
    
    Close #iFileNo
    If Trim(sTemp) <> Trim(sProductKey) Or lTemp <> lRegId Or dTemp <> dExpDate Then
        ValidProductKey = False
        Exit Function
    End If
    ValidProductKey = True

End Function


Public Function FusionUsersLoggedIn() As Integer
    Dim RdoRslt As ADODB.Recordset
    On Error Resume Next
    FusionUsersLoggedIn = 0
    
    sSql = "SELECT COUNT(DISTINCT HOSTNAME) AS 'UsersLoggedIn' FROM master..sysprocesses Where dbid > 0 AND UPPER(program_name)='ES/2000 ERP'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoRslt, ES_FORWARD)
    If bSqlRows Then
        FusionUsersLoggedIn = Val("" & RdoRslt!UsersLoggedIn)
    End If
    Set RdoRslt = Nothing
    If (InStr(1, UCase(Command), "FUSIONROCKS") > 0) Then FusionUsersLoggedIn = 1
End Function

Public Function POMUsersLoggedIn() As Integer
    Dim RdoRslt As ADODB.Recordset
    On Error Resume Next
    POMUsersLoggedIn = 0
    
    sSql = "SELECT COUNT(DISTINCT HOSTNAME) AS 'UsersLoggedIn' FROM master..sysprocesses Where dbid > 0 AND UPPER(program_name)='ESI POINT OF MANUFACTURING'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoRslt, ES_FORWARD)
    If bSqlRows Then
        POMUsersLoggedIn = Val("" & RdoRslt!UsersLoggedIn)
    End If
    Set RdoRslt = Nothing
    If (InStr(1, UCase(Command), "FUSIONROCKS") > 0) Then POMUsersLoggedIn = 1
End Function


Private Function CreateDatePartOfKey() As String
    Dim lJulianDate As Long
    If Not IsDate(dExpDate) Then GetRegistrationInfo
    
    lJulianDate = DateToJulian(Format(dExpDate, "mm/dd/yyyy"))
    CreateDatePartOfKey = Left(ReverseString(Trim(Str(lJulianDate))) + "00000", 5)
End Function


Private Function CreateLicKeyPartOfKey() As String
    Dim lTemp, lTemp2, lTemp3 As Long
    Dim sTemp, sPart1 As String
    If iNumLicenses = 0 Or lRegId = 0 Then GetRegistrationInfo
    
    lTemp2 = 65
    lTemp3 = iNumLicenses
    lTemp = (lTemp2 * lTemp3) + lRegId
    CreateLicKeyPartOfKey = ReverseString(Dec2Hex(lTemp))
End Function



Private Function KeyInfoMatchesDataFile() As Boolean
    Dim iFileNo As Integer
    Dim sTemp1, sTemp2 As String
    
    KeyInfoMatchesDataFile = False
    On Error Resume Next
    
    If Len(sProductKey) = 0 Then GetRegistrationInfo
    
    iFileNo = FreeFile
    Open sAppDir & sLicenseFile For Binary Lock Read Write As #iFileNo
    Get #iFileNo, , LicenseFile
    Close iFileNo
    sTemp1 = Decrypt(LicenseFile.EncProductKey)
    sTemp2 = Decrypt(LicenseFile.EncRegId)
    
    If Trim(UCase(sTemp1)) = Trim(UCase(sProductKey)) And Val(sTemp2) = lRegId Then KeyInfoMatchesDataFile = True
End Function

Public Function RegistrationOk(sMsg As String, Optional POMModule As Boolean = False) As Boolean
   
    If Len(sProductKey) = 0 Then GetRegistrationInfo
    sMsg = ""
            
    If Not ValidProductKey Or Not KeyInfoMatchesDataFile Then
        sMsg = "Invalid License Information has been detected." & vbCrLf & "Do you want to re-register Fusion now?"
        RegistrationOk = False
        Exit Function
    ElseIf IsRegistrationExpired Then
        sMsg = "Fusion Has Expired! Would you like to Register Fusion now?"
        RegistrationOk = False
        Exit Function
    ElseIf (FusionUsersLoggedIn >= LicensedFusionUsersAllowed And Not POMModule) Then
        sMsg = "You have reached your maximun number of allowed users logged into Fusion (" & LicensedFusionUsersAllowed & ")" & vbCrLf & _
        "Would you like to Register more users now?"
        RegistrationOk = False
        Exit Function
    ElseIf (POMUsersLoggedIn >= LicensedPOMUsersAllowed And POMModule) Then
        sMsg = "You have reached your maxiumun number of allowed users logged into POM (" & LicensedPOMUsersAllowed & ")" & vbCrLf & _
          "Would you like to Register more POM users now?"
        RegistrationOk = False
        Exit Function
    End If
    
    If RegDaysLeft < 30 And Not IsRegistrationExpired Then
        sMsg = "You have " & RegDaysLeft & " days left." & vbCrLf & "Would you like to register now?"
        RegistrationOk = True
        Exit Function
    End If
            
    RegistrationOk = True
End Function


Public Function NewProductKeyOk(ByVal sProdKey As String) As Boolean
    Dim sTemp1 As String
    Dim dTemp As Date
    Dim I As Integer
    Dim iLicUsers As Integer
    Dim iPOMUsers As Integer
    
    On Error Resume Next
    Err = 0
    NewProductKeyOk = False
    
    If Len(sCompanyName) = 0 Then GetRegistrationInfo
    
    'Check out first part of key
    sTemp1 = CreateFirstPartOfKey
    If Left(sTemp1, 2) <> Left(sProdKey, 2) Then
        NewProductKeyOk = False
        Exit Function
    End If
    iPOMUsers = LicensedPOMUsersAllowed(sProdKey)
    If iPOMUsers < 0 Or iPOMUsers > 500 Then
        NewProductKeyOk = False
        Exit Function
    End If
            
    'Check out date part of key next
    dTemp = GetExpirationDate(sProdKey)
    If Not IsDate(dTemp) Then
        NewProductKeyOk = False
        Exit Function
    End If
    
    iLicUsers = LicensedFusionUsersAllowed(sProdKey)
    If iLicUsers < 0 Or iLicUsers > 500 Then    'Joel suggested this number on 1/18/2011
        NewProductKeyOk = False
        Exit Function
    End If
    
    ' Check out last part of key
    For I = Len(sProdKey) To 1 Step -1
        If Mid(sProdKey, I, 1) = "-" Then Exit For
    Next I
    sTemp1 = Left(sProdKey, I)
    If CreateLastPartOfKey(sTemp1) <> UCase(Right(sProdKey, 5)) Then
        NewProductKeyOk = False
        Exit Function
    End If
    
    
    If Err > 0 Then NewProductKeyOk = False Else NewProductKeyOk = True
End Function




Public Function RegisterNewKey(ByVal sProdKey As String) As Boolean
    If Not NewProductKeyOk(sProdKey) Then
        RegisterNewKey = False
        Exit Function
    End If
    clsADOCon.ExecuteSQL "Update ComnTable Set ProductKey='" & sProdKey & "'"
        
    sProductKey = sProdKey
    dExpDate = GetExpirationDate(sProdKey)
    
    RegisterNewKey = WriteOutLicenseFile
End Function


