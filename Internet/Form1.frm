VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deskriptor XOR"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8730
   Begin VB.ListBox List5 
      Height          =   6105
      ItemData        =   "Form1.frx":0ECA
      Left            =   6360
      List            =   "Form1.frx":0ECC
      TabIndex        =   21
      Top             =   135
      Width           =   2250
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7740
      TabIndex        =   18
      Top             =   6420
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Get"
      Height          =   495
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   "Get characters of chipertext"
      Top             =   2745
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      Height          =   375
      Left            =   3255
      TabIndex        =   16
      ToolTipText     =   "Undo register character of plainteks"
      Top             =   2280
      Width           =   375
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   4200
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "X"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      ToolTipText     =   "Undo Register character of Chipertext"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Do"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Register the pair-character"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   1230
      ItemData        =   "Form1.frx":0ECE
      Left            =   2520
      List            =   "Form1.frx":0F20
      TabIndex        =   10
      ToolTipText     =   "Alfabet of plaintext character"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "Form1.frx":0F72
      Left            =   1320
      List            =   "Form1.frx":0F74
      TabIndex        =   9
      ToolTipText     =   "Alfabet of Chiper text character "
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save as New File"
      Height          =   495
      Left            =   4395
      TabIndex        =   8
      ToolTipText     =   "Save the descripted text as new file"
      Top             =   4305
      Width           =   1305
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2505
      TabIndex        =   7
      ToolTipText     =   "Plaintext Character"
      Top             =   2265
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "ChiperText Character"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send back"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Send output to Textbox1"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Process"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Changing the character from chipertext"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open File"
      Height          =   495
      Left            =   3975
      TabIndex        =   3
      ToolTipText     =   "Open the encripted file"
      Top             =   2565
      Width           =   990
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "The Result text (Plain Text)"
      Top             =   5040
      Width           =   6015
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":0F76
      Left            =   120
      List            =   "Form1.frx":0F78
      TabIndex        =   1
      ToolTipText     =   "Pair characters"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Chiper text (Enscripted text)"
      Top             =   120
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Panel"
      Height          =   2655
      Left            =   3750
      TabIndex        =   19
      Top             =   2280
      Width           =   2385
      Begin MSComDlg.CommonDialog CoDialog1 
         Left            =   1635
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Undo Send back"
         Height          =   480
         Left            =   1305
         TabIndex        =   20
         Top             =   1440
         Width           =   900
      End
   End
   Begin VB.Label Label2 
      Caption         =   "n ="
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "n = 0"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
      Begin VB.Menu batas1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mMath 
      Caption         =   "Analysis"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*
'* This code is created by Hackerkid (that's me)
'*
'*******************************************************
'*******************************************************
'*
'* There are so many applications using the simple XOR
'* encription as security tool. If u want to know the
'* quality of the XOR encription, u can use this application
'* for testing it. Beside that, by this application
'*  u can learn how to crack your XOR encription and repair yours
'*
'*******************************************************

Dim teks1 As String
Dim teks2 As String
Dim teks3 As String

Private Sub Command1_Click()

    'To open then encripted file, & send the contain to Text1
    Dim tmp As String
    Dim tmp2 As String
    
    With CoDialog1
        .ShowOpen
        If .FileTitle = "" Then Exit Sub
        Open .FileName For Input As 1
            Do While Not EOF(1)
                Input #1, tmp
                tmp2 = tmp2 + tmp
            Loop
        Close #1
    End With

    Text1.Text = tmp2
End Sub

Private Sub Command10_Click()
    'End the application
    End
End Sub

Private Sub Command2_Click()
    Dim data As String
    Dim tmp As String
    Dim i As Integer
    
    'Changing characters of chipertext with characters of plaintext
    'by using data from List1, and then send to Text2.text
    
    If List2.ListCount > 0 Then
        For i = 0 To List2.ListCount - 1
            List1.AddItem List2.List(i) & " --- " & List2.List(i)
        Next
    End If
    
    If List1.ListCount > 0 Then
        data = Text1.Text
        For i = 1 To Len(data)
            tmp = tmp & Ganti(Mid(data, i, 1))
        Next
        Text2.Text = tmp
    End If
End Sub

Private Sub Command3_Click()
    teks3 = Text1.Text
    Text1.Text = Text2.Text
End Sub

Private Sub Command4_Click()
    'to save the plaintext
    On Error Resume Next
    CoDialog1.ShowSave
    Open CoDialog1.FileName For Output As #2
        Print #2, Text2.Text
    Close #2
End Sub

Private Sub Command5_Click()
    'To register the pair characters in List1
    If Text4.Text = "" Then Exit Sub
    List1.AddItem Text3.Text & " --- " & Text4.Text
    Text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Text4.Text <> "" Then Command5.BackColor = vbGreen
End Sub

Private Sub Command6_Click()
    List2.AddItem teks1
    Label1.Caption = "n = " & List2.ListCount
End Sub

Private Sub Command7_Click()
    List3.AddItem teks2
    Label2.Caption = "n = " & List3.ListCount
End Sub

Private Sub Command8_Click()
    Dim i As Integer
    Dim data1 As String
    Dim data2 As String
    
    If List2.List(i) <> "" Then
        If MsgBox("There was exist, Do u want to delete the items in List2 ?", vbOKCancel) = vbOK Then
            List1.Clear
            List2.Clear
            List3.Clear
            List5.Clear
            For i = 0 To List4.ListCount - 1
                List3.AddItem List4.List(i)
            Next
        Else
            Exit Sub
        End If
    End If
    
    data1 = Text1.Text
    For i = 1 To Len(data1)
        data2 = Mid(data1, i, 1)
        If Cek(data2) = False Then List2.AddItem data2
    Next
    Label1.Caption = "n = " & List2.ListCount
    
    For i = 0 To List2.ListCount - 1
        List5.AddItem List2.List(i) & "   = chr ( " & Asc(List2.List(i)) & " )" & "  " & Str(CekJumlah(List2.List(i)))
    Next
End Sub

Private Sub Command9_Click()
        Text1.Text = teks3
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'To add some items in List3,
    'because it's so difficult for me to fill them in List3 by using List3's property
    'I need the items in List4 as my backup
    
    List4.AddItem " "
    For i = 0 To List3.ListCount - 1
        List4.AddItem UCase(List3.List(i))
    Next
    For i = 0 To 9
        List4.AddItem i
    Next
    
    For i = 0 To List4.ListCount - 1
        List3.AddItem List4.List(i)
    Next
    
    Label2.Caption = "n = " & List3.ListCount
    List4.Clear
    For i = 0 To List3.ListCount - 1
        List4.AddItem List3.List(i)
    Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command5.BackColor = vbButtonFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set Form2 = Nothing
End Sub

Private Sub List1_Click()
    Dim data As String
    data = List1.List(List1.ListIndex)
    If MsgBox("Cancel to take that?", vbOKCancel) = vbOK Then
        List2.AddItem Left(data, 1)
        List3.AddItem Right(data, 1)
        List1.RemoveItem List1.ListIndex
    End If
End Sub

Private Sub List2_Click()
    teks1 = List2.List(List2.ListIndex)
    Text3.Text = teks1
    List2.RemoveItem List2.ListIndex
    Label1.Caption = "n = " & List2.ListCount
End Sub

Private Sub List3_Click()
    teks2 = List3.List(List3.ListIndex)
    Text4.Text = teks2
    List3.RemoveItem List3.ListIndex
    Label2.Caption = "n = " & List3.ListCount
End Sub

Private Function Cek(karakter As String) As Boolean
    'is there a same character in List2 ?
    'If nothing, then add the new item character in it.
    
    Dim i As Integer
    Cek = False
    For i = 0 To List2.ListCount - 1
        If karakter = List2.List(i) Then
            Cek = True
            Exit Function
        End If
    Next
End Function

Private Function Ganti(karakter As String) As String
    'To change characters by using items in List1
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If Left(List1.List(i), 1) = karakter Then Ganti = Right(List1.List(i), 1)
    Next
End Function

Private Function CekJumlah(karakter As String) As Integer
    'To know the number of any characters
    Dim i As Integer
    CekJumlah = 0
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) = karakter Then CekJumlah = CekJumlah + 1
    Next
End Function

Private Sub mAbout_Click()
    MsgBox "Created by Hackerkid, firstly"
End Sub

Private Sub mExit_Click()
    End
End Sub

Private Sub mMath_Click()
    Dim i As Integer
    'To show Form2 as a chipertext analyzing form
    With Form2
        For i = 0 To List4.ListCount - 1
            .List2.AddItem List4.List(i) & "   = chr ( " & Asc(List4.List(i)) & " )"
        Next
        .Text1.Text = Text1.Text
        .Panggil
        .Command1.Value = True
        .Show (1)
    End With
End Sub

Private Sub mOpen_Click()
    Command1.Value = True
End Sub

Private Sub Text1_Change()
    Command8.Value = True
End Sub
