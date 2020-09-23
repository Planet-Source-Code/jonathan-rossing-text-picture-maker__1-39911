VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12360
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   17160
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12360
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "Save"
      Height          =   330
      Left            =   3330
      TabIndex        =   18
      Top             =   6165
      Width           =   825
   End
   Begin VB.CommandButton Command9 
      Caption         =   "save"
      Height          =   375
      Left            =   7965
      TabIndex        =   17
      Top             =   11700
      Width           =   960
   End
   Begin VB.TextBox R2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   9765
      Width           =   17025
   End
   Begin VB.CheckBox Check1 
      Caption         =   "DoEvents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   6165
      Width           =   1050
   End
   Begin VB.TextBox R1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4290
      Left            =   4275
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   0
      Width           =   4290
   End
   Begin VB.CommandButton Command8 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command7 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command6 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command5 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   6
      Top             =   6165
      Width           =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "String to text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6885
      TabIndex        =   4
      Top             =   11700
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      MaxLength       =   10
      TabIndex        =   3
      Top             =   11160
      Width           =   6720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pic to text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4230
      TabIndex        =   1
      Top             =   6165
      Width           =   915
   End
   Begin VB.PictureBox P1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   17.5
      ScaleMode       =   4  'Character
      ScaleWidth      =   35
      TabIndex        =   0
      Top             =   0
      Width           =   4260
   End
   Begin VB.PictureBox P2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   45
      ScaleHeight     =   5.001
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   29.924
      TabIndex        =   5
      Top             =   6570
      Width           =   17025
   End
   Begin VB.Label Label3 
      Caption         =   "Text Picture"
      Height          =   240
      Left            =   135
      TabIndex        =   16
      Top             =   9540
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Write a String (max 10 characters)"
      Height          =   240
      Left            =   135
      TabIndex        =   15
      Top             =   10890
      Width           =   2940
   End
   Begin VB.Image Image6 
      Height          =   4200
      Left            =   4365
      Picture         =   "Form1.frx":359E
      Top             =   -45
      Width           =   4200
      Visible         =   0   'False
   End
   Begin VB.Image Image5 
      Height          =   4890
      Left            =   4635
      Picture         =   "Form1.frx":5DD4
      Top             =   -900
      Width           =   3510
      Visible         =   0   'False
   End
   Begin VB.Image Image4 
      Height          =   4875
      Left            =   4950
      Picture         =   "Form1.frx":6476
      Top             =   -675
      Width           =   3045
      Visible         =   0   'False
   End
   Begin VB.Image Image3 
      Height          =   3915
      Left            =   270
      Picture         =   "Form1.frx":68B8
      Top             =   0
      Width           =   7995
      Visible         =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   4755
      Left            =   4680
      Picture         =   "Form1.frx":6DBA
      Top             =   -630
      Width           =   2790
      Visible         =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   5040
      Left            =   2070
      Picture         =   "Form1.frx":71DC
      Top             =   -900
      Width           =   4515
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   2
      Top             =   3015
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim l As Integer

Private Sub Command1_Click()
P1.BackColor = vbWhite

For i = 0 To P1.ScaleHeight
    For j = 1 To P1.ScaleWidth - 1
        If P1.Point(j, i) <> vbWhite Then
            R1.Text = R1.Text & "#"
        ElseIf P1.Point(j, i) = vbWhite Then
            R1.Text = R1.Text & "+"
        End If
        If Check1.Value = Checked Then
        DoEvents
        End If
    Next j
    R1.Text = R1.Text & vbCrLf
    
Next i
End Sub

Private Sub Command10_Click()
    Open App.Path & "\picture.txt" For Append As #1
    Print #1, R1
    Close #1
End Sub

Private Sub Command2_Click()
P2.ScaleMode = 4
R2 = ""

For i = 1 To P2.ScaleHeight
    For j = 1 To c * 5
        If P2.Point(j, i) <> vbBlack Then
            R2.Text = R2.Text & "+"
        ElseIf P2.Point(j, i) = vbBlack Then
            R2.Text = R2.Text & "#"
        End If
    Next j
    R2.Text = R2.Text & vbCrLf
Next i

End Sub



Private Sub Command3_Click()
P1.Picture = Image1
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width

End Sub

Private Sub Command4_Click()
P1.Picture = Image2
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width
End Sub

Private Sub Command5_Click()
P1.Picture = Image3
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width
End Sub

Private Sub Command6_Click()
P1.Picture = Image4
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width
End Sub

Private Sub Command7_Click()
P1.Picture = Image5
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width
End Sub

Private Sub Command8_Click()
P1.Picture = Image6
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
R1.Width = P1.Width
End Sub

Private Sub Command9_Click()
    Open App.Path & "\stringpic.txt" For Append As #1
    Print #1, R2
    Close #1
End Sub

Private Sub Form_Load()
c = 0
P2.BackColor = vbWhite
End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = P1.Point(X, Y)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)


Select Case KeyCode
    Case vbKeyA
        P2.PaintPicture LoadPicture(App.Path & "\a.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyB
        P2.PaintPicture LoadPicture(App.Path & "\b.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyC
        P2.PaintPicture LoadPicture(App.Path & "\c.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyD
        P2.PaintPicture LoadPicture(App.Path & "\d.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyE
        P2.PaintPicture LoadPicture(App.Path & "\e.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyF
        P2.PaintPicture LoadPicture(App.Path & "\f.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyG
        P2.PaintPicture LoadPicture(App.Path & "\g.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyH
        P2.PaintPicture LoadPicture(App.Path & "\h.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyI
        P2.PaintPicture LoadPicture(App.Path & "\i.wmf"), c, 0, 3, 5
        c = c + 1
    Case vbKeyJ
        P2.PaintPicture LoadPicture(App.Path & "\j.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyK
        P2.PaintPicture LoadPicture(App.Path & "\k.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyL
        P2.PaintPicture LoadPicture(App.Path & "\l.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyM
        P2.PaintPicture LoadPicture(App.Path & "\m.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyN
        P2.PaintPicture LoadPicture(App.Path & "\n.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyO
        P2.PaintPicture LoadPicture(App.Path & "\o.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyP
        P2.PaintPicture LoadPicture(App.Path & "\p.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyQ
        P2.PaintPicture LoadPicture(App.Path & "\q.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyR
        P2.PaintPicture LoadPicture(App.Path & "\r.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyS
        P2.PaintPicture LoadPicture(App.Path & "\s.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyT
        P2.PaintPicture LoadPicture(App.Path & "\t.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyY
        P2.PaintPicture LoadPicture(App.Path & "\y.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyU
        P2.PaintPicture LoadPicture(App.Path & "\u.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyW
        P2.PaintPicture LoadPicture(App.Path & "\w.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyV
        P2.PaintPicture LoadPicture(App.Path & "\v.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyY
        P2.PaintPicture LoadPicture(App.Path & "\y.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyX
        P2.PaintPicture LoadPicture(App.Path & "\x.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyZ
        P2.PaintPicture LoadPicture(App.Path & "\z.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeySpace
        P2.PaintPicture LoadPicture(App.Path & "\32.wmf"), c, 0, 3, 5
        c = c + 1
    Case Else
        Beep
    End Select
End Sub
