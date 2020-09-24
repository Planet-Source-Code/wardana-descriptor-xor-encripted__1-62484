VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encriptor"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Descript"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encript"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Too Good for me (Blowfish Method)"
      Height          =   615
      Index           =   2
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "It will spend your live to crack the chipertext"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Almost Good Encriptor"
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "u must more watch the chipertext"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Simple Encriptor"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "It's easy to crack, u can practice your skill"
      Top             =   1200
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2265
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Text            =   "Key"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0ECA
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Blowfish As New clsBlowfish
Public crypto, crypt

Private Sub Command1_Click()
    If Option1(0).Value = True Then Text3.Text = CryptIt(Text1.Text, Text2.Text)
    If Option1(1).Value = True Then Text3.Text = Enkripsi(Text1.Text, Text2.Text, True)
    If Option1(2).Value = True Then Text3.Text = Blowfish.EncryptString(Text1.Text, Text2.Text, False)
End Sub

Private Sub Command2_Click()
    If Option1(0).Value = True Then Text3.Text = CryptIt(Text3.Text, Text2.Text)
    If Option1(1).Value = True Then Text3.Text = Enkripsi(Text3.Text, Text2.Text, False)
    If Option1(2).Value = True Then Text3.Text = Blowfish.DecryptString(Text3.Text, Text2.Text, False)
End Sub

Private Function CryptIt(ToCrypt As String, CryptString As String) As String
    Dim PosS As Long, PosC As Long, TempString As String
    'I'm sorry to The Author of this code but I forget where is this code come from
    TempString = Space$(Len(ToCrypt))
    PosC = 1
    For PosS = 1 To Len(ToCrypt)
        If PosC > Len(CryptString) Then PosC = 1
        Mid(TempString, PosS, 1) = Chr$(Asc(Mid(ToCrypt, PosS, 1)) Xor Asc(Mid(CryptString, PosC, 1)))
        If Asc(Mid(TempString, PosS, 1)) = 0 Then Mid(TempString, PosS, 1) = Mid(ToCrypt, PosS, 1)
        PosC = PosC + 1
    Next PosS
    CryptIt = TempString
End Function

Private Function Enkripsi(data As String, key As String, enscrip As Boolean) As String
    Dim i As Integer
    Dim tmp As String
    Dim tmp2 As String

    If enscrip = True Then
        For i = 1 To Len(data)
            If i Mod 2 = 0 Then
                tmp = Chr((Asc(Mid(data, i, 1)) + 1) Xor Asc(Mid(key, (i Mod Len(key)) + 1, 1)))
            Else
                tmp = Chr((Asc(Mid(data, i, 1)) - 1) Xor Asc(Mid(key, (i Mod Len(key)) + 1, 1)))
            End If
            tmp2 = tmp2 + tmp
        Next
    Else
        For i = 1 To Len(data)
            If i Mod 2 = 0 Then
                tmp = Chr((Asc(Mid(data, i, 1)) Xor Asc(Mid(key, (i Mod Len(key)) + 1, 1))) - 1)
            Else
                tmp = Chr((Asc(Mid(data, i, 1)) Xor Asc(Mid(key, (i Mod Len(key)) + 1, 1))) + 1)
            End If
            tmp2 = tmp2 + tmp
        Next
    End If
    Enkripsi = tmp2
End Function
