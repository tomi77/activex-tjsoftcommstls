VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PasswordGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Alphabets
  alpEnglish = 0
  alpPolish
End Enum

Private mbooSmallChars As Boolean
Private mbooBigChars As Boolean
Private mbooNumbers As Boolean
Private mbooSpecialChars As Boolean
Private mstrSpecialChars As String
Private malpAlphabet As Alphabets

Private Const ENGLISH_SMALL_LETTERS As String = "qwertyuiopasdfghjklzxcvbnm"
Private Const ENGLISH_CAPITAL_LETTERS As String = "QWERTYUIOPASDFGHJKLZXCVBNM"
Private Const POLISH_SMALL_LETTERS As String = "ÍÛπú≥øüÊÒ"
Private Const POLISH_CAPITAL_LETTERS As String = " ”•å£Øè∆—"
Private Const NUMBERS As String = "1234567890"
Private Const DEFAULT_SPECIAL_CHARS As String = "!@#$%^&*()-_=+[{]};:'""\|,<.>/?`~"

Public Property Get UseSmallChars() As Boolean

  UseSmallChars = mbooSmallChars

End Property

Public Property Let UseSmallChars(ByVal NewValue As Boolean)

  mbooSmallChars = NewValue

End Property

Public Property Get UseBigChars() As Boolean

  UseBigChars = mbooBigChars

End Property

Public Property Let UseBigChars(ByVal NewValue As Boolean)

  mbooBigChars = NewValue

End Property

Public Property Get UseNumbers() As Boolean

  UseNumbers = mbooNumbers

End Property

Public Property Let UseNumbers(ByVal NewValue As Boolean)

  mbooNumbers = NewValue

End Property

Public Property Get UseSpecialChars() As Boolean

  UseSpecialChars = mbooSpecialChars

End Property

Public Property Let UseSpecialChars(ByVal NewValue As Boolean)

  mbooSpecialChars = NewValue

End Property

Public Property Get SpecialChars() As String

  SpecialChars = mstrSpecialChars

End Property

Public Property Let SpecialChars(ByVal NewValue As String)

  mstrSpecialChars = NewValue

End Property

Public Property Get Alphabet() As Alphabets

  Alphabet = malpAlphabet

End Property

Public Property Let Alphabet(ByVal NewValue As Alphabets)

  If NewValue >= alpEnglish And NewValue <= alpPolish Then
    malpAlphabet = NewValue
  Else
    malpAlphabet = alpEnglish
  End If

End Property

Private Sub Class_Initialize()

  Randomize
  mbooSmallChars = True
  mbooBigChars = True
  mbooNumbers = True
  mbooSpecialChars = False
  mstrSpecialChars = DEFAULT_SPECIAL_CHARS
  malpAlphabet = alpEnglish

End Sub

Public Function Generate(PassLen As Long)

  Dim strChars As String
  Dim i As Long
  Dim lngCharPos As Long, lngCharsLen As Long
  Dim strPass As String
  
  strChars = ""

  If mbooSmallChars Then
    Select Case Alphabet
    Case alpEnglish
      strChars = strChars & ENGLISH_SMALL_LETTERS
    Case alpPolish
      strChars = strChars & ENGLISH_SMALL_LETTERS & POLISH_SMALL_LETTERS
    End Select
  End If

  If mbooBigChars Then
    Select Case Alphabet
    Case alpEnglish
      strChars = strChars & ENGLISH_CAPITAL_LETTERS
    Case alpPolish
      strChars = strChars & ENGLISH_CAPITAL_LETTERS & POLISH_CAPITAL_LETTERS
    End Select
  End If

  If mbooNumbers Then strChars = strChars & NUMBERS

  If mbooSpecialChars Then strChars = strChars & mstrSpecialChars

  lngCharsLen = Len(strChars)

  For i = 1 To PassLen
    lngCharPos = CInt(Rnd() * (lngCharsLen - 1)) + 1
    strPass = strPass & Mid$(strChars, lngCharPos, 1)
  Next i
  
  Generate = strPass
  
End Function
