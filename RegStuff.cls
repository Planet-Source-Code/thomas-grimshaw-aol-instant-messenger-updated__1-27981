VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReadWriteEasyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Variant
    bInheritHandle As Long
End Type


Public Enum HKeys
    HKEY_CLASSES_ROOT2
    HKEY_CURRENT_USER2
    HKEY_LOCAL_MACHINE2
    HKEY_USERS2
    HKEY_PERFORMANCE_DATA2
    HKEY_CURRENT_CONFIG2
    HKEY_DYN_DATA2
End Enum
Public Enum lDataType
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = (KEY_READ Or KEY_WRITE)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

Dim RootHKey As HKeys, SubDir As String, hKey As Long
Dim OpenRegOk As Boolean

Function GetAllSubDirectories() As Variant
    On Error GoTo handelgetdirvalues
    Dim SubKey_Num As Integer
    Dim SubKey_Name As String
    Dim length As Long
    Dim ReturnArray() As Variant
    If Not OpenRegOk Then Exit Function
    SubKey_Num = 0
    Do
        length = 256
        SubKey_Name = Space(length)
        If RegEnumKey(hKey, SubKey_Num, SubKey_Name, length) <> 0 Then
            Exit Do
        End If
        SubKey_Name = Left(SubKey_Name, InStr(SubKey_Name, Chr$(0)) - 1)
        ReDim Preserve ReturnArray(SubKey_Num) As Variant
        ReturnArray(SubKey_Num) = SubKey_Name
        SubKey_Num = SubKey_Num + 1
    Loop
    GetAllSubDirectories = ReturnArray
    Exit Function
handelgetdirvalues:
    GetAllSubDirectories = Null
    Exit Function
End Function

Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As Variant
    On Error GoTo handelgetavalue
    Dim i As Integer
    Dim SubKey_Value As String, TempStr As String, ReturnArray() As Variant
    Dim length As Long
    Dim RtnVal As Long, value_Type As lDataType
    If Not OpenRegOk Then Exit Function
    RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal 0&, length)
    Select Case RtnVal
        Case 0
        Case 2
            Exit Function
        Case 5
            GetValue = "Access Denied"
            Exit Function
        Case Else
            GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
            Exit Function
    End Select
    SubKey_Value = Space$(length)
    RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal SubKey_Value, length)
    Select Case value_Type
        Case REG_NONE
            SubKey_Value = "Not defined value_type=REG_NONE"
        Case REG_SZ
            SubKey_Value = Left$(SubKey_Value, length - 1)
    Case REG_EXPAND_SZ
        SubKey_Value = Left$(SubKey_Value, length - 1)
    Case REG_BINARY
        SubKey_Value = Left$(SubKey_Value, length)
        If Not ReturnBinStr Then
            TempStr = ""
            For i = 1 To Len(SubKey_Value)
                TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
            Next i
            SubKey_Value = TempStr
        End If
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        SubKey_Value = Left$(SubKey_Value, length)
        If Not ReturnBinStr Then
            TempStr = ""
            For i = 1 To Len(SubKey_Value)
                TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
            Next
            SubKey_Value = TempStr
        End If
    Case REG_DWORD_BIG_ENDIAN
    Case REG_LINK
        SubKey_Value = "Not defined value_type=REG_LINK"
    Case REG_MULTI_SZ
        SubKey_Value = Left$(SubKey_Value, length)
        i = 0
        While Len(SubKey_Value) > 0
            ReDim Preserve ReturnArray(i) As Variant
            ReturnArray(i) = Mid$(SubKey_Value, 1, InStr(1, SubKey_Value, Chr(0)) - 1)
            SubKey_Value = Mid$(SubKey_Value, InStr(1, SubKey_Value, Chr(0)) + 1)
            i = i + 1
        Wend
        GetValue = ReturnArray
        Exit Function
    Case REG_RESOURCE_LIST
        SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"
    Case REG_FULL_RESOURCE_DESCRIPTOR
        SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"
    Case REG_RESOURCE_REQUIREMENTS_LIST
        SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"
    Case Else
        SubKey_Value = "value_type=" & value_Type
    End Select
    GetValue = SubKey_Value
    Exit Function
handelgetavalue:
    GetValue = ""
End Function

Public Property Get RegistryRootKey() As HKeys
    RegistryRootKey = RootHKey
End Property

Public Property Get SubDirectory() As String
    SubDirectory = SubDir
End Property

Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Integer
    On Error GoTo OpenReg
    Dim ReturnVal As Integer
    If RtHKey = 0 Then
        OpenRegistry = False
        OpenRegOk = False
        Exit Function
    End If
    RootHKey = RtHKey
    SubDir = SbDr
    If OpenRegOk Then
        CloseRegistry
        OpenRegOk = False
    End If
    ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_READ_WRITE, hKey)
    If ReturnVal <> 0 Then
        OpenRegistry = False
        Exit Function
    End If
    OpenRegOk = True
    OpenRegistry = True
    Exit Function
OpenReg:
    OpenRegOk = False
    OpenRegistry = False
End Function

Public Function CloseRegistry() As Boolean
    On Error Resume Next
    If RegCloseKey(hKey) <> 0 Then
        CloseRegistry = False
        Exit Function
    End If
    CloseRegistry = True
    OpenRegOk = False
End Function

Private Sub Class_Initialize()
    RootHKey = &H0
    SubDir = ""
    hKey = 0
    OpenRegOk = False
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If RegCloseKey(hKey) <> 0 Then
        Exit Sub
    End If
End Sub
