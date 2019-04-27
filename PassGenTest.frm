VERSION 5.00
Begin VB.Form frmPassGenTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator Test"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAlphabet 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtSpecialChars 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "dodatkowe :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPassLen 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "10"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "cyfry"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "du¿e litery"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "ma³e litery"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generuj"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblPassword 
      Caption         =   "Alfabet :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      Caption         =   "Wygenerowane has³o :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      Caption         =   "D³ugoœæ has³a :"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPassGenTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objPass As New TJSoftCommCtls.PasswordGenerator

Private Sub cboAlphabet_Click()

  objPass.Alphabet = CInt(cboAlphabet.ItemData(cboAlphabet.ListIndex))

End Sub

Private Sub chkOptions_Click(Index As Integer)

  Select Case Index
    Case 0: objPass.UseSmallChars = (chkOptions(Index).Value = CheckBoxConstants.vbChecked)
    Case 1: objPass.useBigChars = (chkOptions(Index).Value = CheckBoxConstants.vbChecked)
    Case 2: objPass.useNumbers = (chkOptions(Index).Value = CheckBoxConstants.vbChecked)
    Case 3: objPass.UseSpecialChars = (chkOptions(Index).Value = CheckBoxConstants.vbChecked)
  End Select

End Sub

Private Sub cmdGenerate_Click()

  If IsNumeric(txtPassLen.Text) Then
    txtPassword.Text = objPass.Generate(CInt(txtPassLen.Text))
  Else
    txtPassword.Text = "<le podana d³ugoœæ has³a !>"
  End If

End Sub

Private Sub Form_Load()

  Dim i As Long

  chkOptions(0).Value = IIf(objPass.UseSmallChars, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
  chkOptions(1).Value = IIf(objPass.useBigChars, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
  chkOptions(2).Value = IIf(objPass.useNumbers, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
  chkOptions(3).Value = IIf(objPass.UseSpecialChars, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
  txtSpecialChars.Text = objPass.SpecialChars
  With cboAlphabet
    .Clear
    .AddItem "Angielski"
    .ItemData(.ListCount - 1) = TJSoftCommCtls.Alphabets.alpEnglish
    .AddItem "Polski"
    .ItemData(.ListCount - 1) = TJSoftCommCtls.Alphabets.alpPolish
    For i = 0 To .ListCount - 1
      If .ItemData(i) = CStr(objPass.Alphabet) Then
        .ListIndex = i
        Exit For
      End If
    Next i
  End With

End Sub

Private Sub txtSpecialChars_Change()

  objPass.SpecialChars = txtSpecialChars.Text

End Sub
