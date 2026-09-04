Attribute VB_Name = "modMain"
Option Base 1
Option Explicit

Public strPasswordComplex(4) As String
Public strApplicationPasswordEncrypted As String
Public strApplicationPasswordCleartext As String
Public strClientOS As String
Public strDomainname As String
Public strDefaultPasswordEncrypted As String
Public strDefaultPasswordClearText As String
Public intPasswordComplexity As Integer
Public intPasswordLength As Integer
Public strPasswordSimple As String
Public strAppPath As String
Public strCurrentUsername As String

Public Enum OSType
    WNT4Server = 0
    WNT4WS = 1
    WNT4TS = 2
    W2KServer = 3
    W2KWS = 4
    W2KTS = 5
    WXPClassic = 6
    WXPDomain = 7
    WXPWorkgroup = 8
    W2003Server = 9
End Enum

Public OS As OSType

Public Type User
    Domain As String
    Username As String
    EncryptedPassword As String
End Type

Public oUsers() As User

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub GetConfig()
    ' Get configuration options
    
    ' Application Password
    strApplicationPasswordEncrypted = GetIniValue("Setup", "Password", "", strAppPath & "setup.ini")
    
    ' Client OS
    strClientOS = GetIniValue("Setup", "Client OS", "", strAppPath & "setup.ini")
    
    ' Domain name
    strDomainname = GetIniValue("Setup", "Domain", "", strAppPath & "setup.ini")
    
    ' Password complexity
    intPasswordComplexity = CInt(GetIniValue("Setup", "Password Complexity", "", strAppPath & "setup.ini"))
    
    ' Simple Password string
    strPasswordSimple = GetIniValue("Setup", "Password Simple String", "", strAppPath & "setup.ini")
    
    ' Complex Password Strings
    strPasswordComplex(1) = GetIniValue("Setup", "Password Complex String 1", "", strAppPath & "setup.ini")
    strPasswordComplex(2) = GetIniValue("Setup", "Password Complex String 2", "", strAppPath & "setup.ini")
    strPasswordComplex(3) = GetIniValue("Setup", "Password Complex String 3", "", strAppPath & "setup.ini")
    strPasswordComplex(4) = GetIniValue("Setup", "Password Complex String 4", "", strAppPath & "setup.ini")
    
    ' Password length
    intPasswordLength = CInt(GetIniValue("Setup", "Password Length", "", strAppPath & "setup.ini"))
    
    ' Default (standard) password
    strDefaultPasswordEncrypted = GetIniValue("Setup", "Standard Password", "", strAppPath & "setup.ini")
    
    ' Current username
    strCurrentUsername = LCase(Environ$("USERNAME"))
    If strCurrentUsername = "" Then strCurrentUsername = "computer_" & LCase(Environ$("COMPUTERNAME"))
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetIniValue
' DateTime  : 03-04-2003 08:21
' Author    : Dave Robinson
' Purpose   : Retrieves a value from an ini file corresponding to the section and key name passed.
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetIniValue(lpSectionName As String, lpKeyName As String, defaultValue As String, inifile As String) As String
    Dim success As Long
    Dim nSize As Long
    Dim Ret As String
    
    On Error GoTo GetIniValue_Error
    
    'call the API with the parameters passed.
    'The return value is the length of the string in ret, including the terminating null. If a default value was passed, and the section or
    'key name are not in the file, that value is returned. If no default value was passed (""), then success will = 0 if not found.
    
    'Pad a string large enough to hold the data.
    Ret = Space$(2048)
    nSize = Len(Ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, defaultValue, Ret, nSize, inifile)
    
    If success Then
        GetIniValue = Left$(Ret, success)
    End If
    
    On Error GoTo 0
    Exit Function
    
GetIniValue_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetIniValue of Module modMain", vbCritical

End Function

Public Sub SaveIniValue(lpSectionName As String, lpKeyName As String, lpValue As String, inifile As String)

'This function saves the passed value to the file,
'under the section and key names specified.
'If the ini file does not exist, it is created.
'If the section does not exist, it is created.
'If the key name does not exist, it is created.
'If the key name exists, it's value is replaced.

   Call WritePrivateProfileString(lpSectionName, lpKeyName, lpValue, inifile)

End Sub

Public Sub DeleteIniSection(lpSectionName As String, inifile As String)

'this call will remove the entire section
'corresponding to lpSectionName. This is
'accomplished by passing vbNullString
'as both the lpKeyName and lpValue parameters.
'For example, assuming that an ini file had:
'  [Colours]
'  Colour1=Red
'  Colour2=Blue
'  Colour3=Green
'
'and this sub was called passing "Colours"
'as lpSectionName, the resulting Colours
'section in the ini file would be deleted.

   Call WritePrivateProfileString(lpSectionName, vbNullString, vbNullString, inifile)

End Sub

Public Sub DeleteIniValue(lpSectionName As String, lpKeyName As String, inifile As String)

'this call will remove the keyname and its
'corresponding value from the section specified
'in lpSectionName. This is accomplished by passing
'vbNullString as the lpValue parameter. For example,
'assuming that an ini file had:
'  [Colours]
'  Colour1=Red
'  Colour2=Blue
'  Colour3=Green
'
'and this sub was called passing "Colour2"
'as lpKeyName, the resulting ini file
'would contain:
'  [Colours]
'  Colour1=Red
'  Colour3=Green

   Call WritePrivateProfileString(lpSectionName, lpKeyName, vbNullString, inifile)

End Sub

Public Function StripNulls(startStrg As String) As String

'take a string separated by nulls,
'split off 1 item, and shorten the string
'so the next item is ready for removal.

'The passed string must have a terminating
'null for this function to work correctly.
'If you remain in a loop, check this first!

   Dim pos As Long
   Dim item As String

   pos = InStr(1, startStrg, Chr$(0))

   If pos Then

      item = Mid$(startStrg, 1, pos - 1)
      startStrg = Mid$(startStrg, pos + 1, Len(startStrg))
      StripNulls = item

   End If

End Function

Public Function Encrypt(strIn As String) As String
    Dim oEncrypt As New clsCryptoFilterBox

    oEncrypt.InBuffer = strIn
    oEncrypt.Encrypt
    Encrypt = oEncrypt.OutBuffer
    
    ' Clean up
    Set oEncrypt = Nothing
End Function

Public Function Decrypt(strIn As String) As String
    Dim oDecrypt As New clsCryptoFilterBox
    
    oDecrypt.InBuffer = strIn
    oDecrypt.Decrypt
    Decrypt = oDecrypt.OutBuffer
    
    ' Clean up
    Set oDecrypt = Nothing
End Function

Public Function NewPassword(Optional Simple As Boolean = True) As String
    Dim i As Integer
    
    For i = 1 To intPasswordLength
        Select Case intPasswordComplexity
            Case 0
                ' Simple
            Case 1
                ' Complex
        End Select
    Next i
End Function
