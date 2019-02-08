VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "首页"
   ClientHeight    =   11100
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   19140
   LinkTopic       =   "Form1"
   ScaleHeight     =   11100
   ScaleWidth      =   19140
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox inputURL 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   11055
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   9735
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   18615
      ExtentX         =   32835
      ExtentY         =   17171
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
   Begin VB.TextBox tmpbar 
      Height          =   615
      Left            =   17640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "执行"
      Height          =   255
      Left            =   17760
      TabIndex        =   2
      Top             =   10680
      Width           =   1095
   End
   Begin VB.ComboBox inputcommand 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   10680
      Width           =   17535
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   15360
      Top             =   240
   End
   Begin MSWinsockLib.Winsock rsv 
      Left            =   16800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   106
   End
   Begin MSWinsockLib.Winsock send 
      Left            =   16080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   105
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   615
      Left            =   14400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1085
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
   Begin VB.Image Image3 
      Height          =   555
      Left            =   13200
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   12240
      Picture         =   "Form1.frx":06E6
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   11400
      Picture         =   "Form1.frx":0E65
      Top             =   -120
      Width           =   585
   End
   Begin VB.Menu mreyWallet 
      Caption         =   "Mrey Wallet"
      Begin VB.Menu newAddress 
         Caption         =   "新地址"
      End
      Begin VB.Menu tuiChu 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu mreyCoin 
      Caption         =   "Mrey Coin"
      Begin VB.Menu sendCoin 
         Caption         =   "发送"
      End
      Begin VB.Menu rsvCoin 
         Caption         =   "接收"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim math
Dim rsa_public As String, rsa_private As String 'RSA加密公钥与私钥
Dim noods(1 To 10) As String '所有节点列表
Public main_nood As String
Public chain_name As String
Dim myWithdrawKey As String




'====================================================================================================================================================================================
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
 
Private m_lOnBits(30)
Private m_l2Power(30)
 Const a As Byte = 20 '密钥
Const b As Byte = 40 '密钥

Private Function StrJiaMi(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource)
'从待加（解）密字符串中取出一个字符
strChar = Mid(strSource, i, 1)
'取字符的低字节和Key1进行异或运算
bLowData = AscB(MidB(strChar, 1, 1)) Xor Key1
'取字符的高字节和K2进行异或运算
bHigData = AscB(MidB(strChar, 2, 1)) Xor key2
'将运算后的数据合成新的字符
If Len(Hex(bLowData)) = 1 Then
strEncrypt = strEncrypt & "0" & Hex(bLowData)
Else
strEncrypt = strEncrypt & Hex(bLowData)
End If
If Len(Hex(bHigData)) = 1 Then
strEncrypt = strEncrypt & "0" & Hex(bHigData)
Else
strEncrypt = strEncrypt & Hex(bHigData)
End If
Next
StrJiaMi = strEncrypt
End Function

Private Function StrJiMi(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource) Step 4
'从待加（解）密字符串中取出一个字符
strChar = Mid(strSource, i, 4)
'取字符的低字节和Key1进行异或运算
bLowData = "&H" & Mid(strChar, 1, 2)
bLowData = bLowData Xor Key1
'取字符的高字节和K2进行异或运算
bHigData = "&H" & Mid(strChar, 3, 2)
bHigData = bHigData Xor key2
'将运算后的数据合成新的字符
strEncrypt = strEncrypt & ChrB(bLowData) & ChrB(bHigData)
Next
StrJiMi = strEncrypt
End Function

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
 
Private Sub md5_FF(a, b, c, d, x, S, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
a = RotateLeft(a, S)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_GG(a, b, c, d, x, S, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
a = RotateLeft(a, S)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_HH(a, b, c, d, x, S, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
a = RotateLeft(a, S)
a = AddUnsigned(a, b)
End Sub
 
Private Sub md5_II(a, b, c, d, x, S, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
a = RotateLeft(a, S)
a = AddUnsigned(a, b)
End Sub
 
Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lblocknumOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
 
Const MODULUS_BITS = 512
Const CONGRUENT_BITS = 448
 
lMessageLength = Len(sMessage)
 
lblocknumOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
ReDim lWordArray(lblocknumOfWords - 1)
 
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
 
lWordArray(lblocknumOfWords - 2) = LShift(lMessageLength, 3)
lWordArray(lblocknumOfWords - 1) = RShift(lMessageLength, 29)
 
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
Private Function Hash(code As String) As String
Hash = MD5(code, 32)
End Function

'======================================================================================================================================================================================
Private Function rsajia(p As Integer, key As String, x As String) As String 'RSA加密



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
Private Function rsajie(p As Integer, key As String, y As String) As String 'RSA解密

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
For i = 2 To b - 1
If (num2 Mod i) = 0 Then
If (num1 Mod i) = 0 Then
hz = False

Exit Function
End If
End If
Next i
hz = True

End Function

Private Sub Command1_Click()
Dim result As String
Dim lastpubkey As String, last_private_key As String
lastpubkey = rsa_public
last_private_key = rsa_private
creatkeys '创建新的rsa密钥
Dim bar As String
bar = StrJiaMi(inputcommand.Text, a, b) + creatdata("nextaddress", Hash(rsa_public))
Debug.Print "sign"
result = creatdata("pubkey", lastpubkey) + creatdata("aut", sign(last_private_key, bar)) + bar

    
If Dir(chain_name, vbDirectory) = "" Then MkDir chain_name
Dim filenum As Long
filenum = 1
While Dir(chain_name + "/" + DataBase("command", filenum) + ".txt") <> ""
    filenum = filenum + 1
Wend
writefile chain_name + "/" + DataBase("command", filenum) + ".txt", result
Debug.Print "checkblock"
If checkblock(chain_name) = "error" Then
    Kill chain_name + "/" + DataBase("command", filenum) + ".txt"
Else:
    tmpbar.Text = checkblock(chain_name)
End If

End Sub

Private Sub Form_Load()
Debug.Print StrJiaMi("null", a, b)
myWithdrawKey = Rnd()
webBrowser.Navigate "https://bing.com"
creatkeys '创建RSA加密密钥
Dim i As Long
For i = 1 To 10
    If Dir(DataBase("nood", i) + ".txt") <> "" Then
        noods(i) = readfile(DataBase("nood", i) + ".txt")
    End If
Next
noods(1) = main_nood
If Len(noods(1)) > 0 Then
    send.RemoteHost = Split(noods(1), ":")(0)
    send.RemotePort = Split(noods(1), ":")(1)
    rsv.Listen
    send.Connect
Else:
    MsgBox "no nood"
End If
End Sub





Private Sub Image1_Click()
inputURL.AddItem inputURL.Text
webBrowser.Navigate inputURL.Text

End Sub

Private Sub Image2_Click()
mycWallet.Show
mycWallet.balance.Caption = "Balance:" + readata(Hash(rsa_public), tmpbar.Text)
mycWallet.address.Caption = "Address:" + Hash(rsa_public)
End Sub

Private Sub Image3_Click()
Dim site As String
site = Split(Replace(inputURL.Text + "/", "://", ""), "/")(0)
Dim bar As String
Dim keynum As Long, key As String
keynum = 1
key = "A"
'visit:site,key,withdraw_hash
While Val(Hash("visit:" + creatdata("site", site) + creatdata("withdrraw_hash", Hash(myWithdrawKey)) + creatdata("key", key))) Mod 1024 <> 1
    keynum = keynum + 1
    key = Hex(keynum)
    DoEvents
Wend
inputcommand.Text = Hash("visit:" + creatdata("site", site) + creatdata("withdrraw_hash", Hash(myWithdrawKey)) + creatdata("key", key))
Call Command1_Click

End Sub

Private Sub newAddress_Click()
creatkeys
End Sub

Private Sub rsv_Connect()
If rsv.State <> sckClosed Then rsv.Close
Dim requestID As Long
rsv.Accept requestID
Dim add_nood As Boolean
add_nood = False
Dim i As Integer
For i = 1 To 10
    If Len(noods(i)) = 0 Then
        add_nood = True
        noods(i) = rsv.RemoteHostIP + ":" + rsv.RemotePort
    End If
Next
End Sub



Private Sub rsv_DataArrival(ByVal bytesTotal As Long)
Dim data As String, info As String
rsv.GetData data
'处理data
info = Split(data, ":")(1)
If InStr(data, "sendIP:") Then
    Dim i As Integer
    For i = 1 To 10
        If Len(noods(i)) = 0 Then noods(i) = info
    Next
ElseIf InStr(data, "send2miner:") Then 'bar(crypto),chain
    Dim bar As String, chain As String
    bar = StrJiMi(readata("bar", info), a, b)
    chain = readata("chain", info)
    Dim filenum As Long
    filenum = 1
    While Dir(chain + "/" + DataBase("command", filenum) + ".txt") <> ""
        filenum = filenum + 1
    Wend
    writefile chain + "/" + DataBase("command", filenum) + ".txt", bar
    If checkblock(chain) = "error" Then Kill chain + "/" + DataBase("command", filenum) + ".txt"
ElseIf InStr(data, "sendblock:") Then 'bar(crypto),chain,num,total
    chain = readata("chain", info)
    If Dir(chain, vbDirectory) = "" Then MkDir chain
    Dim num As Long, total As Long
    num = Val(readata("num", info))
    total = Val(readata("total", info))
    MkDir chain + "2"
    writefile chain + "2/" + DataBase("command", num) + ".txt", StrJiMi(readata("bar", info), a, b)
    If num = total Then '发完了
        '验证区块
        If checkblock(chain + "2") <> "error" Then
            If Dir(chain, vbDirectory) = "" Then
                '第一次下载区块
                Name chain + "2" As chain
            Else:
                If Val(readata("blocknum", checkblock(chain))) > Val(readata("blocknum", checkblock(chain + "2"))) Then
                    '下载的没有主链长
                    RmDir chain + "2"
                Else:
                    RmDir chain
                    Name chain + "2" As chain
                End If
            End If
        End If
    End If
End If
End Sub


Private Sub rsvCoin_Click()
inputcommand.Text = "get:" + creatdata("UnockKey", InputBox("收款指令"))
Call Command1_Click
End Sub

Private Sub send_DataArrival(ByVal bytesTotal As Long)
Dim data As String, info As String
send.GetData data
'处理data
info = Split(data, ":")(1)
If InStr(data, "sendIP:") Then
    Dim i As Integer
    For i = 1 To 10
        If Len(noods(i)) = 0 Then noods(i) = info
    Next
ElseIf InStr(data, "send2miner:") Then 'bar(crypto),chain
    Dim bar As String, chain As String
    bar = StrJiMi(readata("bar", info), a, b)
    chain = readata("chain", info)
    Dim filenum As Long
    filenum = 1
    While Dir(chain + "/" + DataBase("command", filenum) + ".txt") <> ""
        filenum = filenum + 1
    Wend
    writefile chain + "/" + DataBase("command", filenum) + ".txt", bar
    If checkblock(chain) = "error" Then Kill chain + "/" + DataBase("command", filenum) + ".txt"
ElseIf InStr(data, "sendblock:") Then 'bar(crypto),chain,num,total
    chain = readata("chain", info)
    If Dir(chain, vbDirectory) = "" Then MkDir chain
    Dim num As Long, total As Long
    num = Val(readata("num", info))
    total = Val(readata("total", info))
    MkDir chain + "2"
    writefile chain + "2/" + DataBase("command", num) + ".txt", StrJiMi(readata("bar", info), a, b)
    If num = total Then '发完了
        '验证区块
        If checkblock(chain + "2") <> "error" Then
            If Dir(chain, vbDirectory) = "" Then
                '第一次下载区块
                Name chain + "2" As chain
            Else:
                If Val(readata("blocknum", checkblock(chain))) > Val(readata("blocknum", checkblock(chain + "2"))) Then
                    '下载的没有主链长
                    RmDir chain + "2"
                Else:
                    RmDir chain
                    Name chain + "2" As chain
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub sendCoin_Click()

inputcommand.Text = "pay:" + creatdata("seller", InputBox("收款人地址（如果不确定填anyone）:")) + creatdata("cash", InputBox("Mrey Coin 数量:")) + creatdata("UnlockHash", Hash(InputBox("收款指令(如ABC23）："))) + InputBox("注释")
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
Debug.Print "timer start"
Dim i As Long
For i = 1 To 10
    If Len(noods(i)) <> 0 Then
        send.Close
        Debug.Print "timer:close winsock send"
        send.RemoteHost = Split(noods(i), ":")(0)
        send.RemotePort = Split(noods(i), ":")(1)
        '发送数据
        Debug.Print "timer:sending data"
        Dim filenum As Long
        filenum = 0
        While Dir(chain_name + "/" + DataBase("command", filenum + 1) + ".txt") <> ""
            filenum = filenum + 1
        Wend
        Debug.Print "timer:filenum ready"
        send.Connect
        Debug.Print "timer:send has been connect"
        Dim j As Long
        If filenum = 0 Then Exit Sub
        For j = 1 To filenum
            Dim bar As String
            bar = StrJiaMi(readfile(chain_name + "/" + DataBase("command", filenum) + ".txt"), a, b)
            send.SendData "sendblock:" + creatdata("bar", bar) + creatdata("chain", chain_name) + creatdata("num", j) + creatdata("total", filenum)
        Next
    End If
Next

    

End Sub

Private Sub tuiChu_Click()
End
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    If InStr(URL, "a=") Then math = Split(URL, "a=")(1)

End Sub
Private Function DataBase(BaseName As String, blocknum As Long) As String
DataBase = BaseName + "(" & blocknum & ")"
End Function


Private Function sign(prikey As String, code As String) As String
Dim h As String, i As Integer
h = Hash(code)
Dim mychar() As String
ReDim mychar(Len(h) - 1) As String
For i = 1 To Len(h)
    mychar(i - 1) = Mid(h, i, 1)
Next
Dim p As Integer, key As String
p = Val(Split(prikey, ",")(0))
key = Val(Split(prikey, ",")(1))
For i = 1 To Len(h) - 1
    Dim result As String
    If Val(mychar(i)) = 0 Then mychar(i) = 5
    result = result + rsajia(p, key, mychar(i)) + ","
Next
result = Len(h) & "," + result
sign = result
End Function
Private Function checksign(pubkey As String, aut As String, code As String) As Boolean
Dim h As String, mychar() As String
h = Hash(code)
ReDim muchar(Len(h) - 1) As String
Dim i As Long
For i = 1 To Len(h)
    Dim num As Integer
    num = Val(Mid(h, i, 1))
    If num = 0 Then
        mychar(i - 1) = 5
    Else:
        mychar(i - 1) = num
    End If
Next
Dim maxnum As Integer
maxnum = Val(Split(aut, ",")(0))
For i = 1 To maxnum
    Dim p As Integer, key As String
    p = Val(Split(pubkey, ",")(0))
    key = Val(Split(pubkey, ",")(1))
    If rsajie(p, key, Val(Split(aut, ",")(i))) <> mychar(i) Then
        checksign = False
        Exit Function
    End If
Next
checksign = True

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
Private Function check(address As String, nextaddress As String, command As String, tmp0 As String) As String
Dim tmp As String
tmp = tmp0
'读取（创建）block,lasthash,blocknum
Dim block As String, lasthash As String, blocknum As Long
If readata("block", tmp) = "null" Then tmp = tmp + creatdata("block", "") + creatdata("lasthash", "Measure Network") + creatdata("blocknum", 0) + creatdata("comment_num", 0) + creatdata("pool_num", 0) + creatdata("wallet_num", 0)
block = readata("block", tmp)
lasthash = readata("lasthash", tmp)
blocknum = Val(readata("blocknum", tmp))
'检测并创建地址
If readata(address, tmp) = "null" Then tmp = tmp + creatdata(address, 0)
Dim m As Single
m = Val(readata(address, tmp))
Dim comment_num As Long, pool_num As Long, wallet_num As Long
comment_num = Val(readata("comment_num", tmp))
pool_num = Val(readata("pool_num", tmp))
wallet_num = Val(readata("wallet_num", tmp))
If InStr(command, "creat_site:") Then
    '检测数据是否合法
    Dim bar As String
    bar = Replace(command, "creat_site:", "")
    '由于code是加密的所以无需担心
    If InStr(bar, "=") Or InStr(bar, "{") Or InStr(bar, "{") Then
        check = "error"
        Debug.Print "creat_site_data error"
        Exit Function
    Else:
        bar = bar + creatdata("start_time", blocknum) + creatdata("attack_times", 0) + creatdata("protect", 25)
        Dim site As String
        site = readata("site", bar)
        bar = killdata("site", bar)
        tmp = tmp + creatjson(site, bar)
    End If
ElseIf InStr(command, "comment:") Then
    m = m - 50
    If m >= 0 Or Val(Hash(command)) Mod 1024 = 1 Then
        m = m - Val(readata("cash", command))
        If m < 0 Then
            check = "error"
            Debug.Print "comment:no cash"
            Exit Function
        End If
        site = readata("site", command)
        Dim seller As String
        seller = readata("seller", command)
        Dim site_info As String
        site_info = readjson(site, tmp)
        If readata("seller", site_info) = "null" Or readata("seller", site_info) = seller Then
            Dim y As Single
            If readata(seller, tmp) = "null" Then tmp = tmp + creatdata(seller, 0)
            y = Val(readata("cash", seller)) + Val(readata("cash", command))
            tmp = changedata(seller, y, tmp)
            comment_num = comment_num + 1
            Dim i As Long
            i = 1
            While InStr(DataBase("hash", i), command)
                If InStr(readata(DataBase("hash", i), command), "{") Then
                    check = "error"
                    Debug.Print "comment:hash error"
                    Exit Function
                End If
            Wend
            If InStr(readata("text", command), "{") Then
                check = "error"
                Debug.Print "comment:text error"
                Exit Function
            End If
            bar = Replace(command, "comment:", "")
            tmp = tmp + creatjson(DataBase("comment", comment_num), bar + creatdata("ID", Hash(bar)))
        Else:
            check = "error"
            Debug.Print "comment:seller error"
            Exit Function
        End If
    Else:
        check = "error"
        Debug.Print "comment:no cash"
        Exit Function
    End If
ElseIf InStr(command, "reply:") Then
    Dim find_comment As Boolean
    find_comment = False
    For i = 1 To comment_num
        Dim comment_info As String, commentID As String
        commentID = DataBase("comment", i)
        comment_info = readjson(commentID, tmp)
        If readata("ID", comment_info) = readata("ID", command) Then
            find_comment = True
            '检查评论是否被浏览
            Dim j As Long, hashID As String
            Dim find_hash As Boolean
            find_hash = False
            j = 1
            While readata(DataBase("hash", j), comment_info) <> "null"
                hashID = DataBase("hash", j)
                If Hash(readata("key", command)) = readata(hashID, comment_info) Then
                    find_hash = True
                    comment_info = killjson(hashID, comment_info)
                End If
            Wend
            If find_hash = False Then
                check = "error"
                Debug.Print "reply:hash no found"
                Exit Function
            Else:
                '发放奖励
                site = readata("site", comment_info)
                site_info = readjson(site, tmp)
                Dim cost_time As Long, withdraw_key As String, income As Single
                withdraw_key = readata("withdraw_key", command)
                cost_time = Val(readata(Hash(withdraw_key), site_info))
                income = (400 / blocknum) - (20 / (blocknum - Val(readata("strat_time", site_info))))
                site_info = killdata(Hash(withdraw_key), site_info)
                m = income * cost_time + m
                tmp = killjson(site, tmp)
                tmp = killjson(commentID, tmp)
                tmp = tmp + creatjson(commentID, comment_info) + creatjson(site, site_info)
            End If
        End If
    Next
    If find_comment = False Then
        check = "error"
        Debug.Print "reply:comment no found"
        Exit Function
    End If
ElseIf InStr(command, "visit:") Then
    '检查hash是否正确
    If Val(Hash(command)) Mod 1024 = 1 Then
        site = readata("site", command)
        site_info = readjson(site, tmp)
        If site_info = "null" Then
            check = "error"
            Debug.Print "visit:site no found"
            Exit Function
        Else:
            Dim withdraw_hash As String
            withdraw_hash = readata("withdraw_hash", command)
            If readata(withdraw_hash, site_info) = "null" Then
                site_info = site_info + creatdata(withdraw_hash, 1)
            Else:
                cost_time = Val(readata(withdraw_hash, site_info)) + 1
                site_info = changedata(withdraw_hash, cost_time, site_info)
            End If
            tmp = killjson(site, tmp) + creatjson(site, site_info)
        End If
    Else:
        check = "error"
        Debug.Print "visit:hash error"
        Exit Function
    End If
ElseIf InStr(command, "send_protect：") Then
    Dim cash As Single
    cash = Val(readata("cash", command))
    site = readata("site", command)
    site_info = readjson(site, tmp)
    m = m - cash
    If m >= 0 And site_info <> "null" Then
        Dim protect As Single
        protect = Val(readata("protect", site_info)) + cash
        tmp = changedata(site, site_info, tmp)
    Else:
        check = "error"
        Debug.Print "send_protect:no cash or site no found"
        Exit Function
    End If
ElseIf InStr(command, "send_attack:") Then
    cash = Val(readata("cash", command))
    m = m - cash
    If m >= 0 Then
        '检查前面是否有pool_num
        pool_num = pool_num + 1
        site = readata("site", command)
        site_info = readjson(site, tmp)
        If site_info = "null" Then
            check = "error"
            Debug.Print "send.attack site no found"
            Exit Function
        Else:
            protect = Val(readata("protect", site_info))
            bar = Replace(command, "send_attack:", "") + creatdata("risk", protect) + creatdata("start_time", blocknum)
            tmp = tmp + creatjson(DataBase("pool", pool_num), bar)
        End If
    Else:
        check = "error"
        Debug.Print "send.attack no cash"
        Exit Function
    End If
ElseIf InStr(command, "pay:") Then
    cash = Val(readata("cash", command))
    m = m - cash
    If m >= 0 Then
        Dim UnlockHash As String
        UnlockHash = readata("UnlockHash", command)
        If readjson(UnlockHash, tmp) = "null" Then
            tmp = tmp + creatjson(UnlockHash, Replace(command, "pay:", ""))
        Else:
            Debug.Print "pay.UnlockHash has been reged"
            check = "error"
            Exit Function
        End If
    Else:
        check = "error"
        Debug.Print "pay.no cash"
        Exit Function
    End If
ElseIf InStr(command, "get:") Then
    Dim UnlockKey As String
    UnlockKey = readata("UnlockKey", command)
    Dim Wallet As String
    Wallet = readjson(Hash(UnlockKey), tmp)
    If Wallet = "null" Then
        check = " error"
        Debug.Print "get.UnlockHash no found"
        Exit Function
    Else:
        If address = readata("seller", Wallet) Or readata("seller", Wallet) = "anyone" Then
            cash = Val(readata("cash", Wallet))
            If InStr(Wallet, "rsv") Then
                UnlockKey = StrJiMi(UnlockKey, a, b)
                If readata("RemoteAddress", UnlockKey) = address And Val(readata("RemoteCash", UnlockKey)) = cash Then
                    m = m + cash
                    tmp = killjson(Hash(UnlockKey), tmp)
                Else:
                    check = "error"
                    Debug.Print "get.he is a bad guy"
                    Exit Function
                 End If
             Else:
                 m = m + cash
                 tmp = killjson(Hash(UnlockKey), tmp)
             End If
         Else:
             check = "error"
             Debug.Print "get.you are not the owner"
             Exit Function
         End If
    End If
ElseIf InStr(command, "dig:") Then
    cash = Val(readata("cash", command))
    m = m - cash
    If m >= 0 Then
        Dim pool As String, withdraw_address As String
        withdraw_address = readata("withdraw_address", command)
        pool = creatdata("cash", cash) + creatdata("start_time", blocknum) + creatdata("withdraw_address", withdraw_address) + creatdata("risk", "random")
        pool_num = pool_num + 1
        tmp = tmp + creatjson(DataBase("pool", pool_num), pool)
    Else:
        check = "error"
        Debug.Print "dig:no cash"
        Exit Function
    End If
ElseIf InStr(command, "change_hash:") Then
    site = readata("site", command)
    Dim key As String
    key = readata("key", command)
    site_info = readjson(site, tmp)
    Dim next_hash As String
    next_hash = readata("next_hash", command)
    If site_info = "null" Then
        check = "error"
        Debug.Print "change_hash:site no found"
        Exit Function
    Else:
        If readata("change_hash", site_info) = Hash(key) Then
            site_info = changedata("change_hash", next_hash, site_info)
            tmp = killjson(site, tmp)
            tmp = tmp + creatjson(site, site_info)
        Else:
            check = "error"
            Debug.Print "change_hash:key error"
            Exit Function
        End If
    End If
Else:
    check = "error"
    Debug.Print "unkown command"
    Exit Function
End If
'checking mining
If pool_num >= 0 Then
    For i = 1 To pool_num
        Dim poolID As String
        poolID = DataBase("pool", i)
        pool = readjson(poolID, tmp)
        cash = Val(readata("cash", pool))
        withdraw_address = readata("withdraw_address", pool)
        If readata(withdraw_address, tmp) = "null" Then tmp = tmp + creatdata(withdraw_address, 0)
        y = Val(readata(withdraw_address, tmp))
        Dim risk As Long
        If readata("risk", pool) = "random" Then
            If Val(Hash(cash & "")) > Val(Hash(block + lasthash)) * Val("0." + Hash(block)) Then
                'create new block
                blocknum = blocknum + 1
                lasthash = Hash(block + lasthash)
                block = ""
                tmp = killjson(poolID, tmp)
                y = y + cash + (2000 / blocknum)
                tmp = changedata(withdraw_address, y, tmp)
            End If
        Else:
            If blocknum - Val(readata("start_time", pool)) = 2 Then
                Dim fight_result As Long
                fight_result = Val(lasthash) Mod (Val(readata("risk", pool)) + cash)
                If fight_result > cash Then
                    'fail
                    tmp = killjson(poolID, tmp)
                Else:
                    site = readata("site", pool)
                    tmp = killjson(site, tmp)
                    y = y + cash
                    tmp = changedata(withdraw_hash, y, tmp)
                End If
            End If
        End If
    Next
End If
tmp = changedata("blocknum", blocknum, tmp)
tmp = changedata("lasthash", lasthash, tmp)
tmp = changedata("block", block, tmp)
tmp = changedata("pool_num", pool_num, tmp)
tmp = killdata(address, tmp) + creatdata(nextaddress, m)

check = tmp
                
                
            


        
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
rsa_private = n & "," & privatekey
rsa_public = n & "," & publickey
End Sub
Private Function checkblock(chain As String) As String
    Dim i As Long, tmp As String
    tmp = ""
    i = 1
    While Dir(chain + "\" + DataBase("command", i) + ".txt") <> ""
        Dim file As String
        file = readfile(chain + "\" + DataBase("command", i) + ".txt")
        Dim pubkey As String, aut As String, command As String, nextaddress As String, bar As String
        pubkey = readata("pubkey", file)
        aut = readata("aut", file)
        command = StrJiMi(readata("command", file), a, b)
        nextaddress = readata("nextaddress", file)
        bar = command + nextaddress
        If checksign(pubkey, aut, bar) = False Then
            checkblock = "error"
            Debug.Print "sign error"
            Exit Function
        Else:
            tmp = check(Hash(pubkey), nextaddress, command, tmp)
            If tmp = "error" Then
                checkblock = "error"
                Exit Function
            End If
        End If
        i = i + 1
    Wend
    checkblock = tmp
End Function

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
inputURL.Text = URL
Dim S As String, Msg As String, Html As String
With webBrowser.Document
Msg = .body.innerHTML
Msg = Msg + "<script>(function () {var createElement = document.createElement;document.createElement = function (tag) {switch (tag) {case ''script'':console.log(''禁用动态添加脚本，防止广告加载'');break;default:return createElement.apply(this, arguments);}}})();//adblock</script>"
.Clear
.Open
.Write Msg '重写
.Close
If Not InStr(Msg, "//adblock") Then webBrowser.Refresh '刷新
End With
Dim site As String
site = Split(Replace(URL + "/", "://", ""), "/")(0)
If readjson(site, tmpbar.Text) = "null" Then
    inputcommand.Text = "creat_site:" + creatdata("site", site)
    Call Command1_Click
End If
If readjson(site, tmpbar.Text) <> "null" Then
    Dim i As Long, comment_num As Long
    comment_num = Val(readata("comment_num", tmpbar.Text))
    If comment_num > 0 Then
        Dim num As Long
        num = 1
        For i = 1 To comment_num
            Dim comment_info As String
            comment_info = readjson(DataBase("comment", i), tmpbar.Text)
            Dim result As String
            result = creatdata(DataBase("send", num), readata("cash", comment_info) + ":" + StrJiMi(readata("text", tmpbar), a, b))
            num = num + 1
            Dim ads As String
            ads = ads + StrJiMi(readata("text", tmpbar), a, b)
        Next
        If Not InStr(URL, "send") Then
            adsForm.Show
            Open "ads.html" For Output As #1
                Print #1, ads
            Close #1
            adsForm.web.Navigate App.Path + "/ads.html"
            webBrowser.Navigate URL + "?" + result
        End If
    End If
End If

End Sub


