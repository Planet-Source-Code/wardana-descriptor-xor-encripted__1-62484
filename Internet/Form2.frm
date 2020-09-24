VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyzing by using Math Operation"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "1"
      ToolTipText     =   "value in the searching"
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "1"
      ToolTipText     =   "Min Value to search"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Search by using automatic Calculation"
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   2640
      Width           =   5895
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   6960
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   1950
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Send Result"
      Height          =   450
      Left            =   8160
      TabIndex        =   27
      ToolTipText     =   "Back with Send the result to form1"
      Top             =   6705
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
      Height          =   435
      Left            =   9450
      TabIndex        =   26
      ToolTipText     =   "Back without send the result"
      Top             =   6705
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "Automatic Calculation"
      Height          =   2010
      Left            =   4095
      TabIndex        =   21
      Top             =   5160
      Width           =   1965
      Begin VB.CommandButton Command8 
         Caption         =   "Send to Inputbox"
         Height          =   375
         Left            =   165
         TabIndex        =   30
         Top             =   1515
         Width           =   1650
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Root"
         Height          =   195
         Index           =   5
         Left            =   990
         TabIndex        =   29
         Top             =   660
         Width           =   870
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Power"
         Height          =   270
         Index           =   4
         Left            =   975
         TabIndex        =   28
         Top             =   315
         Width           =   840
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Divide"
         Height          =   240
         Index           =   3
         Left            =   105
         TabIndex        =   25
         Top             =   1185
         Width           =   1470
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Time"
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   24
         Top             =   855
         Width           =   1425
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Minus"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   585
         Width           =   1440
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Add"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.TextBox Text4 
      Height          =   1410
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      ToolTipText     =   "Automatic Result Textbox by using median value"
      Top             =   5235
      Width           =   3930
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calculator"
      Height          =   450
      Left            =   4890
      TabIndex        =   19
      Top             =   4605
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Undo to send"
      Height          =   540
      Left            =   1800
      TabIndex        =   18
      Top             =   4560
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send to Inputbox"
      Height          =   540
      Left            =   135
      TabIndex        =   17
      Top             =   4575
      Width           =   1545
   End
   Begin VB.ListBox List2 
      Height          =   4350
      ItemData        =   "Form2.frx":5C12
      Left            =   8400
      List            =   "Form2.frx":5C14
      TabIndex        =   10
      Top             =   105
      Width           =   2040
   End
   Begin VB.ListBox List1 
      Height          =   4350
      ItemData        =   "Form2.frx":5C16
      Left            =   6330
      List            =   "Form2.frx":5C18
      TabIndex        =   9
      Top             =   105
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do"
      Height          =   540
      Left            =   5220
      TabIndex        =   8
      Top             =   2025
      Width           =   795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Root (Sqrt)"
      Height          =   270
      Index           =   5
      Left            =   2550
      TabIndex        =   7
      Top             =   2325
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Power (^)"
      Height          =   330
      Index           =   4
      Left            =   2550
      TabIndex        =   6
      Top             =   1965
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Divide (/)"
      Height          =   270
      Index           =   3
      Left            =   1335
      TabIndex        =   5
      Top             =   2325
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Time (x)"
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   1965
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Minus (-)"
      Height          =   330
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   2295
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Add (+)"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1950
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   1
      Text            =   "2"
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   1770
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Input textbox"
      Top             =   105
      Width           =   5910
   End
   Begin VB.Label Label10 
      Caption         =   "n                    ="
      Height          =   240
      Left            =   6330
      TabIndex        =   39
      Top             =   6915
      Width           =   1590
   End
   Begin VB.Label Label9 
      Caption         =   "Length           ="
      Height          =   255
      Left            =   6315
      TabIndex        =   38
      Top             =   6585
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Median Value = 77"
      Height          =   330
      Left            =   8400
      TabIndex        =   33
      Top             =   6240
      Width           =   2040
   End
   Begin VB.Label Label7 
      Caption         =   "Max. Value    = 122"
      Height          =   330
      Left            =   8400
      TabIndex        =   32
      Top             =   5880
      Width           =   2040
   End
   Begin VB.Label Label6 
      Caption         =   "Length Value = 90"
      Height          =   330
      Left            =   8400
      TabIndex        =   16
      Top             =   5520
      Width           =   2040
   End
   Begin VB.Label Label5 
      Caption         =   "Min. Value     = 32"
      Height          =   330
      Left            =   8415
      TabIndex        =   15
      Top             =   5160
      Width           =   2040
   End
   Begin VB.Label Label4 
      Caption         =   "Median Value ="
      Height          =   330
      Left            =   6285
      TabIndex        =   14
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Max. Value    ="
      Height          =   330
      Left            =   6315
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Length Value ="
      Height          =   330
      Left            =   6330
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Min. Value     ="
      Height          =   330
      Left            =   6345
      TabIndex        =   11
      Top             =   5160
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nilai As Integer
Dim nilai2 As Integer
Dim nilai3 As Integer
Dim backup As String

Private Sub Command1_Click()
    'Base on the value of option1 (nilai), we manipulate each character in Chipertext
    'by using math operation and send the result to text3
    Dim data1 As String
    Dim data2 As String
    Dim i As Integer
    On Error Resume Next
    For i = 1 To Len(Text1.Text)
        Select Case nilai
            Case 0
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) + Val(Text2.Text))
            Case 1
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) - Val(Text2.Text))
            Case 2
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) * Val(Text2.Text))
            Case 3
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) / Val(Text2.Text))
            Case 4
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) ^ Val(Text2.Text))
            Case 5
                data1 = Chr(Asc(Mid(Text1.Text, i, 1)) ^ (1 / Val(Text2.Text)))
        End Select
        data2 = data2 + data1
    Next
    Text3.Text = data2
End Sub

Private Sub Command2_Click()
    'Change the chipertext with result of math operation
    backup = Text1.Text     'As our backup if we want to change back
    Text1.Text = Text3.Text
    Panggil
End Sub

Private Sub Command3_Click()
    'Change back the chipertext
    If backup <> "" Then Text1.Text = backup
End Sub

Private Sub Command4_Click()
    'Call Calculator of Windows
   Shell Environ("Windir") & "\System32\calc.exe"
End Sub

Private Sub Command5_Click()
    Form2.Hide
    Form1.Show
End Sub

Private Sub Command6_Click()
    Form1.Text1.Text = Text3.Text
    Form2.Hide
End Sub

Private Function Cek(karakter As String) As Boolean
    'To know if there is a same character in List3
    Dim i As Integer
    Cek = False
    For i = 0 To List3.ListCount - 1
        If karakter = List3.List(i) Then
            Cek = True
            Exit Function
        End If
    Next
End Function

Private Function CekJumlah(karakter As String) As Integer
    'To know the number of character
    Dim i As Integer
    CekJumlah = 0
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) = karakter Then CekJumlah = CekJumlah + 1
    Next
End Function

Private Sub Command8_Click()
    backup = Text1.Text
    Text1.Text = Text4.Text
    Panggil
End Sub

Private Sub Command9_Click()
    'The automatic process to calculate the value of character
    'base on the value of option2 (nilai2)
    
    On Error Resume Next
    Dim j As Integer
    Dim tmp1 As String
    Dim tmp2 As String
    
        nilai3 = nilai3 + Val(Text6.Text)
        tmp1 = ""
        tmp2 = ""
        Select Case nilai2
            Case 0
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) + nilai3)
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            Case 1
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) - nilai3)
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            Case 2
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) * nilai3)
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            Case 3
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) / nilai3)
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            Case 4
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) ^ nilai3)
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            Case 5
                For j = 1 To Len(Text1.Text)
                    tmp1 = Chr(Asc(Mid(Text1.Text, j, 1)) ^ (1 / nilai3))
                    tmp2 = tmp2 + tmp1
                Next j
                Text7.Text = Str(nilai3)
            End Select
            Text4.Text = tmp2
End Sub

Private Sub Form_Load()
    nilai = 0
    nilai2 = 0
End Sub

Private Sub Option1_Click(Index As Integer)
    'Send the value of Index to nilai
    Select Case Index
        Case 0
            nilai = 0
        Case 1
            nilai = 1
        Case 2
            nilai = 2
        Case 3
            nilai = 3
        Case 4
            nilai = 4
        Case 5
            nilai = 5
    End Select
    'Call Command1_Click (it's the same)
    Command1.Value = True
End Sub

Private Sub Option2_Click(Index As Integer)
    'Send the value of option2 to nilai2
    Select Case Index
        Case 0
            nilai2 = 0
        Case 1
            nilai2 = 1
        Case 2
            nilai2 = 2
        Case 3
            nilai2 = 3
        Case 4
            nilai2 = 4
        Case 5
            nilai2 = 5
    End Select
    nilai3 = 0
    Command9.Value = True
End Sub

Private Sub Text1_Change()
    Panggil
End Sub

Public Sub Panggil()
    'Prosedur to fill items in List1
    'which are character informations of Chipertext
    Dim i As Integer
    Dim data1 As String
    Dim data2 As String
    Dim min As Integer
    Dim max As Integer
    
    If Text1.Text = "" Then Exit Sub
    List1.Clear
    List3.Clear
    
    data1 = Text1.Text
    For i = 1 To Len(data1)
        data2 = Mid(data1, i, 1)
        
        If Cek(data2) = False Then List3.AddItem data2
    Next
    
    For i = 0 To List3.ListCount - 1
        List1.AddItem List3.List(i) & "   = chr ( " & Asc(List3.List(i)) & " )" & "  " & Str(CekJumlah(List3.List(i)))
    Next
    
    min = Asc(List3.List(0))
    max = Asc(List3.List(List3.ListCount - 1))
    Label1.Caption = "Min. Value     = " & Str(min)
    Label2.Caption = "Length Value = " & Str(max - min)
    Label3.Caption = "Max Value     = " & Str(max)
    Label4.Caption = "Median Value = " & Str((max + min) / 2 + min)
    Label9.Caption = "Length           = " & Len(Text1.Text)
    Label10.Caption = "n                    = " & List1.ListCount
End Sub
