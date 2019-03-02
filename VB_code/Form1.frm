VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Mrey"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   10110
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox inputURL 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   8055
   End
   Begin MSWinsockLib.Winsock winsock 
      Left            =   9480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   9480
      Top             =   2280
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6855
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   12091
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
 
Private m_lOnBits(30)
Private m_l2Power(30)
 Dim rsa_public As String, rsa_private As String
 Dim math As String
 Public timer As Boolean
Public winsockResult As String
Dim tmp As String, adslist As String
Private Function LShift(lValue, iShiftBits)
If iShiftBits = 0 Then
LShift = lValue
Exit Function
ElseIf iShiftBits = 31 Then
If lValue And 1 Then
LShift = &H80000000
Else
LShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If
 
If (lValue And m_l2Power(31 - iShiftBits)) Then
LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
Else
LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End If
End Function
 
Private Function RShift(lValue, iShiftBits)
If iShiftBits = 0 Then
RShift = lValue
Exit Function
ElseIf iShiftBits = 31 Then
If lValue And &H80000000 Then
RShift = 1
Else
RShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If
 
RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
 
If (lValue And &H80000000) Then
RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End If
End Function
 
Private Function RotateLeft(lValue, iShiftBits)
RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
 
Private Function AddUnsigned(lX, lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult
 
lX8 = lX And &H80000000
lY8 = lY And &H80000000
lX4 = lX And &H40000000
lY4 = lY And &H40000000
 
lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
If lX4 And lY4 Then
lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
ElseIf lX4 Or lY4 Then
If lResult And &H40000000 Then
lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
Else
lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
End If
Else
lResult = lResult Xor lX8 Xor lY8
End If
 
AddUnsigned = lResult
End Function
 
Private Function md5_F(x, y, z)
md5_F = (x And y) Or ((Not x) And z)
End Function
 
Private Function md5_G(x, y, z)
md5_G = (x And z) Or (y And (Not z))
End Function
 
Private Function md5_H(x, y, z)
md5_H = (x Xor y Xor z)
End Function
 
Private Function md5_I(x, y, z)
md5_I = (y Xor (x Or (Not z)))
End Function
 
Private Sub md5_FF(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_GG(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_HH(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_II(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
 
Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
 
Const MODULUS_BITS = 512
Const CONGRUENT_BITS = 448
 
lMessageLength = Len(sMessage)
 
lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
ReDim lWordArray(lNumberOfWords - 1)
 
lBytePosition = 0
lByteCount = 0
Do Until lByteCount >= lMessageLength
lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
lByteCount = lByteCount + 1
Loop
 
lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
 
lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
 
lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
 
ConvertToWordArray = lWordArray
End Function
 
Private Function WordToHex(lValue)
Dim lByte
Dim lCount
 
For lCount = 0 To 3
lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
Next
End Function
 
Public Function MD5(sMessage, stype)
m_lOnBits(0) = CLng(1)
m_lOnBits(1) = CLng(3)
m_lOnBits(2) = CLng(7)
m_lOnBits(3) = CLng(15)
m_lOnBits(4) = CLng(31)
m_lOnBits(5) = CLng(63)
m_lOnBits(6) = CLng(127)
m_lOnBits(7) = CLng(255)
m_lOnBits(8) = CLng(511)
m_lOnBits(9) = CLng(1023)
m_lOnBits(10) = CLng(2047)
m_lOnBits(11) = CLng(4095)
m_lOnBits(12) = CLng(8191)
m_lOnBits(13) = CLng(16383)
m_lOnBits(14) = CLng(32767)
m_lOnBits(15) = CLng(65535)
m_lOnBits(16) = CLng(131071)
m_lOnBits(17) = CLng(262143)
m_lOnBits(18) = CLng(524287)
m_lOnBits(19) = CLng(1048575)
m_lOnBits(20) = CLng(2097151)
m_lOnBits(21) = CLng(4194303)
m_lOnBits(22) = CLng(8388607)
m_lOnBits(23) = CLng(16777215)
m_lOnBits(24) = CLng(33554431)
m_lOnBits(25) = CLng(67108863)
m_lOnBits(26) = CLng(134217727)
m_lOnBits(27) = CLng(268435455)
m_lOnBits(28) = CLng(536870911)
m_lOnBits(29) = CLng(1073741823)
m_lOnBits(30) = CLng(2147483647)
 
m_l2Power(0) = CLng(1)
m_l2Power(1) = CLng(2)
m_l2Power(2) = CLng(4)
m_l2Power(3) = CLng(8)
m_l2Power(4) = CLng(16)
m_l2Power(5) = CLng(32)
m_l2Power(6) = CLng(64)
m_l2Power(7) = CLng(128)
m_l2Power(8) = CLng(256)
m_l2Power(9) = CLng(512)
m_l2Power(10) = CLng(1024)
m_l2Power(11) = CLng(2048)
m_l2Power(12) = CLng(4096)
m_l2Power(13) = CLng(8192)
m_l2Power(14) = CLng(16384)
m_l2Power(15) = CLng(32768)
m_l2Power(16) = CLng(65536)
m_l2Power(17) = CLng(131072)
m_l2Power(18) = CLng(262144)
m_l2Power(19) = CLng(524288)
m_l2Power(20) = CLng(1048576)
m_l2Power(21) = CLng(2097152)
m_l2Power(22) = CLng(4194304)
m_l2Power(23) = CLng(8388608)
m_l2Power(24) = CLng(16777216)
m_l2Power(25) = CLng(33554432)
m_l2Power(26) = CLng(67108864)
m_l2Power(27) = CLng(134217728)
m_l2Power(28) = CLng(268435456)
m_l2Power(29) = CLng(536870912)
m_l2Power(30) = CLng(1073741824)
 
 
Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d
 
Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21
 
x = ConvertToWordArray(sMessage)
 
a = &H67452301
b = &HEFCDAB89
c = &H98BADCFE
d = &H10325476
 
For k = 0 To UBound(x) Step 16
AA = a
BB = b
CC = c
DD = d
 
md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
 
md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
md5_GG d, a, b, c, x(k + 10), S22, &H2441453
md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
 
md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
 
md5_II a, b, c, d, x(k + 0), S41, &HF4292244
md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
md5_II c, d, a, b, x(k + 6), S43, &HA3014314
md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
 
a = AddUnsigned(a, AA)
b = AddUnsigned(b, BB)
c = AddUnsigned(c, CC)
d = AddUnsigned(d, DD)
Next
 
If stype = 32 Then
MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
Else
MD5 = LCase(WordToHex(b) & WordToHex(c))
End If
End Function

Private Function Hash(code) As String
Hash = MD5(code, 16)
End Function





Private Sub Command1_Click()
If InStr(inputURL.Text, "seed") Then writefile "noods.txt", readfile("noods.txt") + readata("seed", inputURL.Text) + ";"
Dim noods As String
noods = readfile("noods.txt")
Dim num As Long
num = countString(noods, ";")
If num > 0 Then
    Dim i As Long
    For i = 1 To num
        Dim host As String
        host = Split(noods, ";")(i - 1)
        winsock.Close
        winsock.RemoteHost = Split(host, ":")(0)
        winsock.RemotePort = Val(Split(host, ":")(1))
        On Error Resume Next
        winsock.Connect
        timer = False
        Timer1.Enabled = True
        Dim filehash As String
        filehash = readata("filehash", inputURL.Text)
        While timer = False
            DoEvents
        Wend
        Timer1.Enabled = False
        If Dir(filehash + ".html") = "" Then
            wb.Navigate2 readata("server", inputURL.Text)
            Dim xmlHTTP1
            Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
            xmlHTTP1.Open "get", "http://www.qq.com", True
            xmlHTTP1.send
            While xmlHTTP1.ReadyState <> 4
                DoEvents
            Wend
            writefile filehash + ".html", xmlHTTP1.responseText
            Set xmlHTTP1 = Nothing
        End If
    Next
Else:
    MsgBox "No peer"
End If



        
End Sub

Private Sub Form_Load()
math = "null"
creatkeys
End Sub
Private Function countString(base As String, findString As String) As Long
countString = Len(base) - Len(Replace(base, findString, ""))
End Function
Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    If InStr(URL, "a=") Then
        math = Split(URL, "a=")(1)
    ElseIf InStr(URL, "#error") Then
        math = URL
    ElseIf InStr(URL, "#") Then
        math = URL
    End If

End Sub
Private Function DataBase(BaseName As String, blocknum As Long) As String
DataBase = BaseName + "(" & blocknum & ")"
End Function
Private Function creatjson(json As String, data As String) As String
creatjson = json + "{" + data + "}"
End Function
Private Function readjson(json As String, tmp As String) As String
readjson = bstr(tmp, json + "{", "}")
End Function
Private Function killjson(json As String, tmp As String) As String
Dim data As String
data = readjson(json, tmp)
killjson = Replace(tmp, creatjson(json, data), "")
End Function
Private Function hz(num1 As Integer, num2 As Integer) As Boolean '互质
Dim i As Integer

For i = 2 To num1 - 1
If (num1 Mod i) = 0 Then
If (num2 Mod i) = 0 Then
hz = False
Exit Function

End If
End If
Next i
For i = 2 To num2 - 1
If (num2 Mod i) = 0 Then
If (num1 Mod i) = 0 Then
hz = False

Exit Function
End If
End If
Next i
hz = True

End Function
Public Sub creatkeys()
Dim p As Long, q As Long, n As Integer, m As Integer
p = 29
q = 31
n = p * q
m = (p - 1) * (q - 1)
Dim publickey As Integer, privatekey As Integer
privatekey = 10
While hz(privatekey, m) = False
    privatekey = privatekey + 1
Wend
publickey = 10
While publickey * privatekey Mod m <> 1
    publickey = publickey + 1
Wend
rsa_private = "P" & n & "K" & privatekey & "X"
rsa_public = "P" & n & "K" & publickey & "X"
End Sub
Private Function rsajia(p As String, key As String, x As String) As String 'RSA加密



Open "rsa.html" For Output As #1
Print #1, "<script>"
Print #1, "var a = Math.pow(" & x & "," & key & ")%" & p & ";"
Print #1, "window.location.href = ""http://127.0.0.1/?a=""+a;"
Print #1, "</script>"
Close #1
wb.Navigate App.Path + "/rsa.html"
While math = "null"
    DoEvents
Wend
rsajia = math
End Function
Private Function rsajie(p As String, key As String, y As String) As String 'RSA解密

Open "rsa.html" For Output As #1
Print #1, "<script>"
Print #1, "var a = Math.pow(" & y & "," & key & ")%" & p & ";"
Print #1, "window.location.href = ""http://127.0.0.1/?a=""+a;"
Print #1, "</script>"
Close #1
wb.Navigate App.Path + "/rsa.html"
While math = "null"
    DoEvents
Wend
rsajie = math
End Function

Private Function bstr(code, str1, str2)
    On Error GoTo e
    bstr = Split(Split(code, str1)(1), str2)(0)
    Exit Function
e:
    
    bstr = "null"
End Function

Private Function creatdata(BaseName, data) As String
    creatdata = BaseName + "=" & data & ";"
End Function
Private Function readata(BaseName, base)
    On Error GoTo e
    readata = bstr(base, BaseName + "=", ";")
    Exit Function
e:
    readata = "null"
End Function
Private Function readfile(filename As String) '读取文件内容
On Error Resume Next
Open filename For Input As #1
Dim b As String

   Do While Not EOF(1)
       Input #1, b
       readfile = readfile + b
   Loop
   Close #1
End Function
Public Sub writefile(filename As String, data As String) '输出文件
    Open filename For Output As #1
        Print #1, data
    Close #1
        
End Sub
Private Function changedata(BaseName, newdata, base) As String
    Dim oldstr As String, newstr As String
    On Error GoTo e
    oldstr = creatdata(BaseName, readata(BaseName, base))
    newstr = creatdata(BaseName, newdata)
    On Error GoTo e
    changedata = Replace(base, oldstr, newstr)
    Exit Function
e:
    changedata = base
    
End Function
Private Function killdata(BaseName, base) As String
    On Error GoTo e
    killdata = Replace(base, BaseName + "=" + readata(BaseName, base) + ";", "")
    Exit Function
e:
    killdata = base
    
End Function
Public Function StrToHex(ByVal strS As String) As String
'将字符串转换为16进制
    Dim abytS() As Byte
    Dim bytTemp As Byte
    Dim strTemp As String
    Dim lLocation As Long
    abytS = StrConv(strS, vbFromUnicode)
    For lLocation = 0 To UBound(abytS)
        bytTemp = abytS(lLocation)
        strTemp = Hex(bytTemp)
        strTemp = Right("00" & strTemp, 2)
        StrToHex = StrToHex & strTemp
    Next lLocation
    StrToHex = StrToHex
End Function
 
Public Function HexToStr(str As String) As String
'将16进制转换为字符串
    Dim rst() As Byte
    Dim i As Long, j As Long, strlong As Long
    strlong = Len(str)
    ReDim rst(strlong \ 2)
    For i = 0 To strlong - 1 Step 2
        Dim tmp As Long
        rst(i / 2) = Val("&H" & Mid(str, i + 1, 2))
    Next
    HexToStr = StrConv(rst, vbUnicode)
End Function

Private Sub winsock_ConnectionRequest(ByVal requestID As Long)
If winsock.State <> sckClosed Then winsock.Close
winsock.Accept requestID
End Sub

Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
Dim data As String
winsock.GetData data
Dim helpTimes As Long
helpTimes = Val(readata(winsock.RemoteHost, tmp))
If InStr(data, "askHash-") Then
    Dim filename As String
    filename = readata("filename", data)
    If Dir(filename) <> "" And helpTimes > -2 Then
        winsock.SendData "reply-" + creatdata("txt", StrToHex(readfile(filename)))
        helpTimes = helpTimes - 1
        tmp = changedata(winsock.RemoteHost, helpTimes, tmp)
    End If
ElseIf InStr(data, "reply-") Then
    filename = readata("filename", data)
    Dim txt As String
    txt = readata("txt", data)
    If filename = "noods.txt" Then
        writefile "noods.txt", readfile("noods.txt") + HexToStr(txt)
        helpTimes = helpTimes + 1
        tmp = changedata(winsock.RemoteHost, helpTimes, tmp)
    ElseIf InStr(filename, ".html") Then
        Dim filehash As String
        filehash = Split(filename, ".html")
        txt = Replace(HexToStr(txt), vbCrLf, "")
        If Hash(Split(txt, "<!-- Mrey Mining Module -->")(0)) = filehash Then
            txt = "<script>var check=""error"";</script>" + txt + "<script>window.location.href=""#check:""+check;</script>"
            writefile "Temp_" + filename, txt
            math = "null"
            wb.Navigate App.Path + "/Temp_" + filename
            While math = "null"
                DoEvents
            Wend
            If InStr(math, "check:") And Not InStr(math, "error") Then
                Dim tmpNum As Long
                tmpNum = Val(Split(math, "check:")(1))
                If Dir(filename) = "" Then
                    FileCopy "Temp_" + filename, filename
                    helpTimes = helpTimes + 1
                    wb.Navigate App.Path + "/Temp_" + filename
                Else:
                    If Val(Split(math, "check:")(1)) > tmpNum Then
                        FileCopy "Temp_" + filename, filename
                        helpTimes = helpTimes + 1
                    Else:
                        helpTimes = helpTimes + 1
                        wb.Navigate App.Path + "/" + filename
                    End If
                End If
            Else:
                helpTimes = helpTimes - 1
            End If
        Else:
            helpTimes = helpTimes - 1
        End If
    Else:
        helpTimes = helpTimes - 1
    End If
    tmp = changedata(winsock.RemoteHost, helpTimes, tmp)
End If
            
End Sub


