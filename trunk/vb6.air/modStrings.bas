Attribute VB_Name = "modString"
      Option Explicit
      Public Enum KnownCodePage
          CP_UNKNOWN = -1
          CP_ACP = 0
          CP_OEMCP = 1
          CP_MACCP = 2
          CP_THREAD_ACP = 3
          CP_SYMBOL = 42
      '   ARABIC
          CP_AWIN = 101   ' Bidi Windows codepage
          CP_709 = 102    ' MS-DOS Arabic Support CP 709
          CP_720 = 103    ' MS-DOS Arabic Support CP 720
          CP_A708 = 104   ' ASMO 708
          CP_A449 = 105   ' ASMO 449+
          CP_TARB = 106   ' MS Transparent Arabic
          CP_NAE = 107    ' Nafitha Enhanced Arabic Char Set
          CP_V4 = 108     ' Nafitha v 4.0
          CP_MA2 = 109    ' Mussaed Al Arabi (MA/2) CP 786
          CP_I864 = 110   ' IBM Arabic Supplement CP 864
          CP_A437 = 111   ' Ansi 437 codepage
          CP_AMAC = 112   ' Macintosh Code Page
      '   HEBREW
          CP_HWIN = 201   ' Bidi Windows codepage
          CP_862I = 202   ' IBM Hebrew Supplement CP 862
          CP_7BIT = 203   ' IBM Hebrew Supplement CP 862 Folded
          CP_ISO = 204    ' ISO Hebrew 8859-8 Character Set
          CP_H437 = 205   ' Ansi 437 codepage
          CP_HMAC = 206   ' Macintosh Code Page
      '   CODE PAGES
          CP_OEM_437 = 437
          CP_ARABICDOS = 708
          CP_DOS720 = 720
          CP_DOS737 = 737
          CP_DOS775 = 775
          CP_IBM850 = 850
          CP_IBM852 = 852
          CP_DOS861 = 861
          CP_DOS862 = 862
          CP_IBM866 = 866
          CP_DOS869 = 869
          CP_THAI = 874
          CP_EBCDIC = 875
          CP_JAPAN = 932
          CP_CHINA = 936
          CP_KOREA = 949
          CP_TAIWAN = 950
      '   UNICODE
          CP_UNICODELITTLE = 1200
          CP_UNICODEBIG = 1201
      '   CODE PAGES
          CP_EASTEUROPE = 1250
          CP_RUSSIAN = 1251
          CP_WESTEUROPE = 1252
          CP_GREEK = 1253
          CP_TURKISH = 1254
          CP_HEBREW = 1255
          CP_ARABIC = 1256
          CP_BALTIC = 1257
          CP_VIETNAMESE = 1258
      '   KOREAN
          CP_JOHAB = 1361
      '   MAC
          CP_MAC_ROMAN = 10000
          CP_MAC_JAPAN = 10001
          CP_MAC_ARABIC = 10004
          CP_MAC_GREEK = 10006
          CP_MAC_CYRILLIC = 10007
          CP_MAC_LATIN2 = 10029
          CP_MAC_TURKISH = 10081
      '   CODE PAGES
          CP_CHINESECNS = 20000
          CP_CHINESEETEN = 20002
          CP_IA5WEST = 20105
          CP_IA5GERMAN = 20106
          CP_IA5SWEDISH = 20107
          CP_IA5NORWEGIAN = 20108
          CP_ASCII = 20127
          CP_RUSSIANKOI8R = 20866
          CP_RUSSIANKOI8U = 21866
          CP_ISOLATIN1 = 28591
          CP_ISOEASTEUROPE = 28592
          CP_ISOTURKISH = 28593
          CP_ISOBALTIC = 28594
          CP_ISORUSSIAN = 28595
          CP_ISOARABIC = 28596
          CP_ISOGREEK = 28597
          CP_ISOHEBREW = 28598
          CP_ISOTURKISH2 = 28599
          CP_ISOLATIN9 = 28605
          CP_HEBREWLOG = 38598
          CP_USER = 50000
          CP_AUTOALL = 50001
          CP_JAPANNHK = 50220
          CP_JAPANESC = 50221
          CP_JAPANISO = 50222
          CP_KOREAISO = 50225
          CP_TAIWANISO = 50227
          CP_CHINAISO = 50229
          CP_AUTOJAPAN = 50932
          CP_AUTOCHINA = 50936
          CP_AUTOKOREA = 50949
          CP_AUTOTAIWAN = 50950
          CP_AUTORUSSIAN = 51251
          CP_AUTOGREEK = 51253
          CP_AUTOARABIC = 51256
          CP_JAPANEUC = 51932
          CP_CHINAEUC = 51936
          CP_KOREAEUC = 51949
          CP_TAIWANEUC = 51950
          CP_CHINAHZ = 52936
          CP_GB18030 = 54936
      '   UNICODE
          CP_UTF7 = 65000
          CP_UTF8 = 65001
      End Enum
      ' Flags
      Public Const MB_PRECOMPOSED = &H1
      Public Const MB_COMPOSITE = &H2
      Public Const MB_USEGLYPHCHARS = &H4
      Public Const MB_ERR_INVALID_CHARS = &H8
      Public Const WC_DEFAULTCHECK = &H100                ' check for default char
      Public Const WC_COMPOSITECHECK = &H200              ' convert composite to precomposed
      Public Const WC_DISCARDNS = &H10                    ' discard non-spacing chars
      Public Const WC_SEPCHARS = &H20                     ' generate separate chars
      Public Const WC_DEFAULTCHAR = &H40                  ' replace with default char
      Public Declare Function GetACP Lib "kernel32" () As Long
      Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, _
      ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
      ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
      Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
      ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
      ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
      lpUsedDefaultChar As Long) As Long
      
      Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
      
      Public Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
      
      
      Public Function ANSItoUTF16(ByRef Text() As Byte, Optional ByVal cPage As KnownCodePage = CP_UNKNOWN, _
                                  Optional lFlags As Long) As Byte()
          Static tmpArr() As Byte, textStr As String
          Dim tmpLen As Long, textLen As Long, A As Long
          If (Not Text) = True Then Exit Function
          ' set code page to a valid one
          If cPage = CP_UNKNOWN Then cPage = GetACP
          If cPage = CP_ACP Or cPage = CP_WESTEUROPE Then
              textLen = UBound(Text)
              tmpLen = textLen + textLen + 1
              If (Not tmpArr) = True Then ReDim Preserve tmpArr(tmpLen)
              If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen)
              For A = 0 To UBound(Text)
                  tmpArr(A + A) = Text(A)
              Next A
          Else
              textStr = CStr(Text) & "|"
              textLen = LenB(textStr)
              tmpLen = textLen + textLen
              ReDim Preserve tmpArr(tmpLen + 1)
              ' get the new string to tmpArr
              tmpLen = MultiByteToWideChar(CLng(cPage), lFlags, ByVal StrPtr(textStr), -1, _
                                           ByVal VarPtr(tmpArr(0)), tmpLen)
              If tmpLen = 0 Then Exit Function
              tmpLen = tmpLen + tmpLen - 5
              'If tmpArr(tmpLen - 1) = 0 And tmpArr(tmpLen) = 0 Then tmpLen = tmpLen - 2
              If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen)
          End If
          ' return the result
          ANSItoUTF16 = tmpArr
      End Function
      
      
      Public Function UTF16toANSI(ByRef Text() As Byte, Optional ByVal cPage As KnownCodePage = CP_UNKNOWN, _
                                  Optional lFlags As Long) As Byte()
          Static tmpArr() As Byte
          Dim tmpLen As Long, textLen As Long, A As Long
          If (Not Text) = True Then Exit Function
          ' set code page to a valid one
          If cPage = CP_UNKNOWN Then cPage = GetACP
          If cPage = CP_ACP Or cPage = CP_WESTEUROPE Then
              textLen = UBound(Text)
              tmpLen = (textLen + 1) \ 2 - 1
              If (Not tmpArr) = True Then ReDim Preserve tmpArr(tmpLen)
              If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen)
              For A = 0 To tmpLen
                  tmpArr(A) = Text(A + A)
              Next A
          Else
              textLen = (UBound(Text) + 1) \ 2
              ' at maximum ANSI can be four bytes per character in new Chinese encoding GB18030–2000
              tmpLen = textLen + textLen + textLen + textLen + 1
              ReDim Preserve tmpArr(tmpLen - 1)
              ' get the new string to tmpArr
              tmpLen = WideCharToMultiByte(CLng(cPage), lFlags, ByVal VarPtr(Text(0)), textLen, ByVal VarPtr(tmpArr(0)), _
                                           tmpLen, ByVal 0&, ByVal 0&)
              If tmpLen = 0 Then Exit Function
              ' a hopeless try to correct a weird error?
              ReDim Preserve tmpArr(tmpLen - 1)
          End If
          ' return the result
          UTF16toANSI = tmpArr
      End Function


Public Function OemToCharS(sOutput As String)
   Dim outputstr As String
   outputstr = Space$(Len(sOutput))
   OemToChar sOutput, outputstr
   OemToCharS = outputstr
End Function

Public Function ToOEM(sourcestring As String)
    Dim deststring As String  ' получаемая строка
    Dim code As Long
    
    deststring = Space$(Len(sourcestring)) 'получаем перекодированную строку
    code = CharToOem(sourcestring, deststring)
    ToOEM = deststring
End Function
