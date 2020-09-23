Attribute VB_Name = "Module1"
'Option Explicit
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Global mainkeyhandle As Long, strbuf4$, Scode, done, result



Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Dim bfound
Dim hKey As Long


    #If Win32 Then
        
        Public Const HKEY_CLASSES_ROOT = &H80000000
        Public Const HKEY_CURRENT_USER = &H80000001
        Public Const HKEY_LOCAL_MACHINE = &H80000002
        Public Const HKEY_USERS = &H80000003
        Public Const KEY_ALL_ACCESS = &H3F
        Public Const REG_OPTION_NON_VOLATILE = 0&
        Public Const REG_CREATED_NEW_KEY = &H1
        Public Const REG_OPENED_EXISTING_KEY = &H2
        Public Const ERROR_SUCCESS = 0&

    #End If


Type SECURITY_ATTRIBUTES
    
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
    End Type


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Public Declare Func
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public Const EM_GETLINECOUNT = &HBA

Public Const GWL_STYLE As Long = -16&
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8

Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003

Public Const TVIF_STATE As Long = &H8

'treeview styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_CHECKBOXES = &H100
Public Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Public Const TVIS_CUT        As Long = &H4
Public Const TVIS_BOLD       As Long = &H10
Public Const TVIS_CHECK      As Long = &H3000
Public Const TVIS_CHECKED    As Long = &H2000
Public Const TVIS_UNCHECKED  As Long = &H1000

Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TVITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Sub FormatBold(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
    
   With frmSetInfo.InfoEditor
        If (IsNull(.SelBold) = True) Or (.SelBold = False) Or (vntForce = True) Then
            .SelBold = True
          
        ElseIf .SelBold = True Then
            .SelBold = False
            
        End If
    End With
End Sub

Sub FormatItalic(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
    With frmSetInfo.InfoEditor
        If (IsNull(.SelItalic) = True) Or (.SelItalic = False) Or (vntForce = True) Then
            .SelItalic = True
           
        ElseIf .SelItalic = True Then
           
            .SelItalic = False
           
        End If
    End With
End Sub

Sub FormatUnderline(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
   With frmSetInfo.InfoEditor
        If (IsNull(.SelUnderline) = True) Or _
            (.SelUnderline = False) Or (vntForce = True) Then
                       
            .SelUnderline = True
           
        ElseIf .SelUnderline = True Then
            
            .SelUnderline = False
            
        End If
    End With
End Sub

Sub FormatAlign(mdiFrm As Form, intIndex As Integer)
    With frmSetInfo
        Select Case intIndex
            Case 0
                frmSetInfo.InfoEditor.SelAlignment = rtfLeft
            Case 1
                 frmSetInfo.InfoEditor.SelAlignment = rtfCenter
            Case 2
                 frmSetInfo.InfoEditor.SelAlignment = rtfRight
        End Select
    End With
End Sub

Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String

Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

'check for lngStartPosition ad lngEndPosition

If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)

lngLastFontColor& = -1 'no color

   For lngCurText& = lngStartPosition& To lngEndPosition&
       rtbRichTextBox.SelStart = lngCurText&
       rtbRichTextBox.SelLength = 1
   
          If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
             intLastAlignment% = rtbRichTextBox.SelAlignment
              
                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"
                   Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"
                   Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
                End Select
                
          End If
   
          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 strHTML$ = strHTML$ & "<b>"
               Else
                 strHTML$ = strHTML$ & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 strHTML$ = strHTML$ & "<u>"
               Else
                 strHTML$ = strHTML$ & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 strHTML$ = strHTML$ & "<i>"
               Else
                 strHTML$ = strHTML$ & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 strHTML$ = strHTML$ & "<s>"
               Else
                 strHTML$ = strHTML$ & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """>"
         End If

         If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If
 
     strHTML$ = strHTML$ + rtbRichTextBox.SelText

   Next lngCurText&

RichToHTML = strHTML$

End Function
Public Function GetURLSource2(scodez As String) As String
frmSignOn.Winsock1.Close

frmSignOn.Winsock1.RemoteHost = "64.12.163.213"
frmSignOn.Winsock1.RemotePort = 5190
done = False
Scode = scodez
frmSignOn.Winsock1.Connect


Do Until done = True

DoEvents
Loop
GetURLSource2 = result


End Function



Public Sub MakeBold()
Dim TVI As TVITEM
Dim r As Long
Dim hitemTV As Long
Dim hwndTV As Long
Dim tmpNode As Node

   hwndTV = frmBuddyList.tvwBuddies.hwnd
   hitemTV = SendMessageLong(hwndTV, TVM_GETNEXTITEM, TVGN_CARET, 0&)
  'if a valid handle get and set the  'item's state attributes
   If hitemTV > 0 Then
      With TVI
         .hItem = hitemTV
         .mask = TVIF_STATE
         .stateMask = TVIS_BOLD
         r = SendMessageAny(hwndTV, TVM_GETITEM, 0&, TVI)
         'flip the bold mask state
         Select Case .state And TVIS_BOLD
           Case TVIS_BOLD
             .state = 0
           Case Else
             .state = TVIS_BOLD
         End Select
      End With
      r = SendMessageAny(hwndTV, TVM_SETITEM, 0&, TVI)
   End If
   Set frmBuddyList.tvwBuddies.SelectedItem = tmpNode
End Sub





'I think this was put in by The Hobo (original programmer of GetInfo routine).
'
'Haven't deleted it just in case
'
'Public Function Dates()
'Dim i As Long
'For i = 1 To 31
'Dates = i
'Next
'End Function
Public Function bSetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    
    'on error resume next
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue, CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    bSetRegValue = (lResult = ERROR_SUCCESS)
    
End Function



Public Function bGetRegValue(ByVal sKey As String, ByVal sSubKey As String) As String
    hKey = HKEY_LOCAL_MACHINE
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim szBuffer2 As String
    Dim lBuffSize2 As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    
    lIndex = 0
    lResult = RegOpenKeyEx(hKey, sKey, 0, 1, phkResult)


    Do While lResult = ERROR_SUCCESS And Not (bfound)
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)


        If (lResult = ERROR_SUCCESS) Then
            sCompKey = Left(szBuffer, lBuffSize)


            If (sCompKey = sSubKey) Then
                bGetRegValue = Left(szBuffer2, lBuffSize2 - 1)
            End If
        End If
        lIndex = lIndex + 1
        
    Loop
    RegCloseKey phkResult
End Function


Function CreateKey(subKey As String)


hKey = HKEY_LOCAL_MACHINE

Call ParseKey(subKey, Format(mainkeyhandle))

If mainkeyhandle Then
   rtn = RegCreateKey(mainkeyhandle, subKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub

Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function
Function htmltorich2(strings)
Call HTMLToRich(Format$(strings), rtfform.RichTextBox1)
htmltorich2 = rtfform.RichTextBox1.TextRTF

End Function
Sub HTMLToRich(strHTML As String, rtbRichTextBox As RichTextLib.RichTextBox)
Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long, lngLastFontSize As Long
Dim lngChar As Long, strTag As String, lngSpot As Long, strChar As String
Dim lngAlign As Long, strBuf As String, strBuf2 As String, lngBuf As Long, strBuf3 As String

Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

'set default values
strLastFont$ = rtbRichTextBox.Font.Name
lngLastFontColor& = -1

'clear richtextbox
rtbRichTextBox.Text = ""



   'Loop through string. If finds an HTML string
   For lngChar& = 1 To Len(strHTML$)
   DoEvents
      strChar$ = Mid$(strHTML$, lngChar&, 1)
      
        If strChar$ = "<" Then
           lngSpot& = InStr(lngChar& + 1, strHTML$, ">")
              If lngSpot& Then
              
                 strTag$ = LCase$(Mid$(strHTML$, lngChar& + 1, lngSpot& - lngChar& - 1))
                 
                   If strTag$ = "b" Then
                      blnBold = True
                   ElseIf strTag$ = "/b" Then
                      blnBold = False
                   ElseIf strTag$ = "u" Then
                      blnUnderline = True
                   ElseIf strTag$ = "/u" Then
                      blnUnderline = False
                   ElseIf strTag$ = "i" Then
                      blnItalic = True
                   ElseIf strTag$ = "/i" Then
                      blnItalic = False
                   ElseIf strTag$ = "s" Then
                      blnStrikeThru = True
                   ElseIf strTag$ = "/s" Then
                      blnStrikeThru = False
                   ElseIf Left$(strTag$, 8) = "p align=" Then
                      strBuf$ = Right$(strTag$, Len(strTag$) - 8)
                      strBuf3$ = ""
                      
                         For lngBuf& = 1 To Len(strBuf$)
                              strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                              If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                         Next lngBuf&
                         
                         Select Case strBuf3$
                             Case "left":   lngAlign& = AlignLeft
                             Case "right":  lngAlign& = AlignRight
                             Case "center": lngAlign& = AlignCenter
                         End Select
                         
                   ElseIf Left$(strTag$, 5) = "font " Or Left$(strTag$, 5) = "FONT " Then
                   toggle = False
                      strbuf4$ = Right$(strTag$, Len(strTag$) - 5)
                        For xerses = 1 To Len(strbuf4$)
                        If Mid$(strbuf4$, xerses, 1) = """" Then toggle = Not toggle
                        innit = False
                        If Mid$(strbuf4$, xerses, 1) = " " And toggle = False Then innit = True Else innit = False
                        If innit = True Or xerses = Len(strbuf4$) Then
                        If xerses = Len(strbuf4$) Then khg$ = khg$ + Mid$(strbuf4$, xerses, 1)
                        
                        strBuf$ = khg$
                        khg$ = ""
                         Select Case Left$(strBuf$, InStr(strBuf$, "=") - 1)
                            
                            Case "color":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                  For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" And strBuf2$ <> "#" Then strBuf3$ = strBuf3$ & strBuf2$
                                  Next lngBuf&
                               Red = Hex2Dec(Left$(strBuf3$, 2))
                               Green = Hex2Dec(Mid$(strBuf3$, 3, 2))
                               Blue = Hex2Dec(Right$(strBuf3$, 2))
                               lngLastFontColor& = Red + Green * 256 + Blue * 256 * 256
                               
                               
                            Beep
                            Case "COLOR":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                  For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" And strBuf2$ <> "#" Then strBuf3$ = strBuf3$ & strBuf2$
                                  Next lngBuf&
                               lngLastFontColor& = Hex2Dec(Left$(strBuf3$, 6))
                            Beep
                            Case "face":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                  For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                                  Next lngBuf&
                               strLastFont$ = strBuf3$
                            Case "FACE":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                  For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                                  Next lngBuf&
                               strLastFont$ = strBuf3$
                            
                            Case "size":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                   For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                                   Next lngBuf&
                                   
                                   Select Case strBuf3$
                                      Case "1": lngLastFontSize& = 4
                                      Case "2": lngLastFontSize& = 8
                                      Case "3": lngLastFontSize& = 10
                                      Case "4": lngLastFontSize& = 14
                                      Case "5": lngLastFontSize& = 18
                                      Case "6": lngLastFontSize& = 20
                                      Case "7": lngLastFontSize& = 72
                                   End Select
                              Case "SIZE":
                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                   For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" Then strBuf3$ = strBuf3$ & strBuf2$
                                   Next lngBuf&
                                   
                                   Select Case strBuf3$
                                      Case "1": lngLastFontSize& = 4
                                      Case "2": lngLastFontSize& = 8
                                      Case "3": lngLastFontSize& = 10
                                      Case "4": lngLastFontSize& = 14
                                      Case "5": lngLastFontSize& = 18
                                      Case "6": lngLastFontSize& = 20
                                      Case "7": lngLastFontSize& = 72
                                   End Select
                         End Select
                         Else
                       khg$ = khg$ + Mid$(strbuf4$, xerses, 1)
                       End If
                       Next xerses
                       
                   End If
                   
                 'skip over html tag
                 lngChar& = lngSpot&
              End If 'for: If lngSpot& Then
        Else
           'set character with curretn artributes.
           rtbRichTextBox.SelStart = Len(rtbRichTextBox.Text)
           rtbRichTextBox.SelLength = 0
           rtbRichTextBox.SelText = strChar$
           rtbRichTextBox.SelStart = Len(rtbRichTextBox.Text) - 1
           rtbRichTextBox.SelLength = 1
           rtbRichTextBox.SelBold = blnBold
           rtbRichTextBox.SelUnderline = blnUnderline
           rtbRichTextBox.SelItalic = blnItalic
           rtbRichTextBox.SelStrikeThru = blnStrikeThru
           rtbRichTextBox.SelFontName = strLastFont$
           rtbRichTextBox.SelFontSize = lngLastFontSize&
           rtbRichTextBox.SelAlignment = lngAlign&
           rtbRichTextBox.SelColor = lngLastFontColor&
        End If 'for: If rtbRichTextBox.SelText = "<" Then
     

      
    Next lngChar&

rtbRichTextBox.SelStart = Len(rtbRichTextBox.Text) - 1
rtbRichTextBox.SelLength = 1
End Sub
Function ErrorMsg(lErrorCode As Long) As String
    Dim GetErrorMsg
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function
Function HexToDecimal(ByVal strHex As String) As Long

'This function is required by the function 'HTMLToRich'

'this function converts any hexidecimal color value
'(e.g. "0000FF" = Blue) to decimal color value.

Dim lngDecimal As Long, strCharHex As String, lngColor As Long
Dim lngChar As Long
If Left$(strHex$, 1) = "=" Then strHex$ = Right$(strHex$, Len(strHex$) - 1)
If Left$(strHex$, 1) = "#" Then strHex$ = Right$(strHex$, 6)
  
strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
If Left$(strHex$, 1) = "=" Then strHex$ = Right$(strHex$, Len(strHex$) - 1)
  For lngChar& = Len(strHex$) To 1 Step -1
    strCharHex$ = Mid$(UCase$(strHex$), lngChar&, 1)
    
       Select Case strCharHex$
          Case 0 To 9
             lngDecimal& = CLng(strCharHex$)
          Case Else 'A,B,C,D,E,F
             lngDecimal& = CLng(Chr$((Asc(strCharHex$) - 17))) + 10
       End Select
       
    lngColor& = lngColor& + lngDecimal& * 16 ^ (Len(strHex$) - lngChar&)
  Next lngChar&
  
HexToDecimal = lngColor&

End Function
