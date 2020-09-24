VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circulatory System (c) Shamansoft Incorporation"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Main.frx":08CA
   ScaleHeight     =   8625
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7170
      Left            =   5640
      Picture         =   "Main.frx":1D1FE
      ScaleHeight     =   7170
      ScaleWidth      =   5715
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   5715
      Begin VB.PictureBox picBack 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1440
         Picture         =   "Main.frx":1F10F
         ScaleHeight     =   465
         ScaleWidth      =   840
         TabIndex        =   35
         Top             =   4080
         Width           =   840
      End
      Begin VB.PictureBox picRight 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   4080
         Picture         =   "Main.frx":205A9
         ScaleHeight     =   465
         ScaleWidth      =   1215
         TabIndex        =   33
         Top             =   4080
         Width           =   1215
      End
      Begin VB.PictureBox picLeft 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   2760
         Picture         =   "Main.frx":22377
         ScaleHeight     =   465
         ScaleWidth      =   1215
         TabIndex        =   32
         Top             =   4080
         Width           =   1215
      End
      Begin VB.PictureBox picPlay 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   480
         Picture         =   "Main.frx":24145
         ScaleHeight     =   465
         ScaleWidth      =   840
         TabIndex        =   30
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox txtDefine 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   4800
         Width           =   5175
      End
      Begin VB.PictureBox picHeart 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3600
         Left            =   495
         Picture         =   "Main.frx":255DF
         ScaleHeight     =   3600
         ScaleWidth      =   4800
         TabIndex        =   27
         ToolTipText     =   "Right-click to copy picture"
         Top             =   360
         Width           =   4800
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   240
            Top             =   840
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   240
            Top             =   360
         End
      End
      Begin VB.PictureBox picStop 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   480
         Picture         =   "Main.frx":36CD3
         ScaleHeight     =   465
         ScaleWidth      =   840
         TabIndex        =   31
         Top             =   4080
         Width           =   840
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   1215
      Left            =   360
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2143
      _Version        =   393217
      TextRTF         =   $"Main.frx":3816D
   End
   Begin VB.PictureBox picShowLabel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   6240
      MouseIcon       =   "Main.frx":381EF
      Picture         =   "Main.frx":384F9
      ScaleHeight     =   465
      ScaleWidth      =   2565
      TabIndex        =   23
      Top             =   6240
      Width           =   2565
   End
   Begin VB.PictureBox picTopic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5730
      Left            =   480
      Picture         =   "Main.frx":3C3B7
      ScaleHeight     =   5730
      ScaleWidth      =   4875
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5235
         ItemData        =   "Main.frx":3DD5E
         Left            =   240
         List            =   "Main.frx":3DDA1
         TabIndex        =   28
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.PictureBox picShowSystem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   6240
      MouseIcon       =   "Main.frx":3DF68
      Picture         =   "Main.frx":3E272
      ScaleHeight     =   465
      ScaleWidth      =   2565
      TabIndex        =   24
      Top             =   6240
      Width           =   2565
   End
   Begin VB.PictureBox Picture3 
      Height          =   3135
      Left            =   2280
      ScaleHeight     =   3075
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox pSystem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   1
         Left            =   240
         Picture         =   "Main.frx":42130
         ScaleHeight     =   465
         ScaleWidth      =   2565
         TabIndex        =   22
         Top             =   2400
         Width           =   2565
      End
      Begin VB.PictureBox pSystem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":45FEE
         ScaleHeight     =   465
         ScaleWidth      =   2565
         TabIndex        =   21
         Top             =   1920
         Width           =   2565
      End
      Begin VB.PictureBox pLabel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   1
         Left            =   240
         Picture         =   "Main.frx":49EAC
         ScaleHeight     =   465
         ScaleWidth      =   2565
         TabIndex        =   20
         Top             =   1320
         Width           =   2565
      End
      Begin VB.PictureBox pLabel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":4DD6A
         ScaleHeight     =   465
         ScaleWidth      =   2565
         TabIndex        =   19
         Top             =   840
         Width           =   2565
      End
      Begin VB.PictureBox picClick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   4680
         Picture         =   "Main.frx":51C28
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   13
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picClick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   4080
         Picture         =   "Main.frx":52B72
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   12
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picClick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   3480
         Picture         =   "Main.frx":53ABC
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   11
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   2040
         Picture         =   "Main.frx":54A06
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   10
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   1440
         Picture         =   "Main.frx":55950
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   9
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   840
         Picture         =   "Main.frx":5689A
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   8
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picClick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   2880
         Picture         =   "Main.frx":577E4
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   7
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":5872E
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   6
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   9240
      Picture         =   "Main.frx":59678
      ScaleHeight     =   1695
      ScaleWidth      =   1980
      TabIndex        =   3
      Top             =   6240
      Width           =   1980
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   720
         MouseIcon       =   "Main.frx":64586
         Picture         =   "Main.frx":64890
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   16
         Top             =   720
         Width           =   510
      End
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   1320
         MouseIcon       =   "Main.frx":657DA
         Picture         =   "Main.frx":65AE4
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   15
         Top             =   720
         Width           =   510
      End
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   720
         MouseIcon       =   "Main.frx":66A2E
         Picture         =   "Main.frx":66D38
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   14
         Top             =   120
         Width           =   510
      End
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   120
         MouseIcon       =   "Main.frx":67C82
         Picture         =   "Main.frx":67F8C
         ScaleHeight     =   555
         ScaleWidth      =   510
         TabIndex        =   4
         Top             =   720
         Width           =   510
      End
   End
   Begin VB.PictureBox picView 
      BackColor       =   &H00FFFFFF&
      Height          =   4770
      Left            =   6120
      ScaleHeight     =   4710
      ScaleWidth      =   5160
      TabIndex        =   0
      Top             =   1320
      Width           =   5220
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox picSystem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   15780
         Left            =   -720
         Picture         =   "Main.frx":68ED6
         ScaleHeight     =   15750
         ScaleWidth      =   7200
         TabIndex        =   1
         ToolTipText     =   "Click on the heart to view the parts of the heart"
         Top             =   0
         Width           =   7230
         Begin VB.Label lblHeart 
            BackStyle       =   0  'Transparent
            Height          =   1095
            Index           =   0
            Left            =   3000
            MouseIcon       =   "Main.frx":7792E
            MousePointer    =   99  'Custom
            TabIndex        =   17
            ToolTipText     =   "Click here to view the parts of the  heart"
            Top             =   3120
            Width           =   975
         End
      End
      Begin VB.PictureBox picLabel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   15780
         Left            =   -720
         Picture         =   "Main.frx":77C38
         ScaleHeight     =   15750
         ScaleWidth      =   7200
         TabIndex        =   2
         ToolTipText     =   "Click on the heart to view the parts of the heart"
         Top             =   0
         Width           =   7230
         Begin VB.Label lblHeart 
            BackStyle       =   0  'Transparent
            Height          =   1095
            Index           =   1
            Left            =   3000
            MouseIcon       =   "Main.frx":8CCE9
            MousePointer    =   99  'Custom
            TabIndex        =   18
            ToolTipText     =   "Click here to view the parts of the  heart"
            Top             =   3120
            Width           =   975
         End
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy picture to clipboard"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ButtonClick As String
Dim DisplayType As String

Dim Pos As String

Dim Beat As Integer
Dim Tick As Boolean

Private Sub lblHeart_Click(Index As Integer)

Beat = 1
picDisplay.Visible = True
picTopic.Visible = True

picPlay.Visible = True
picStop.Visible = False

txtDefine.Text = ""

End Sub

Private Sub List1_Click()

rtf.LoadFile App.Path & "\data\" & LTrim(Str(List1.ListIndex)) & ".txt", rtfText
picHeart.Picture = LoadPicture(App.Path & "\pictures\parts\" & LTrim(Str(List1.ListIndex)) & ".jpg")

Timer2.Enabled = False
Timer3.Enabled = False

picPlay.Visible = True
picStop.Visible = False

End Sub

Private Sub mnuCopy_Click()

Clipboard.SetData picHeart
MsgBox "Picture has been copied to clipboard.", , "Message"

End Sub

Private Sub picBack_Click()

picDisplay.Visible = False
picTopic.Visible = False

Timer2.Enabled = False
Timer3.Enabled = False

End Sub

Private Sub picControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        ButtonClick = "LEFT"
    Case 1
        ButtonClick = "UP"
    Case 2
        ButtonClick = "RIGHT"
    Case 3
        ButtonClick = "DOWN"
End Select

Timer1.Enabled = True

End Sub

Private Sub picControl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer1.Enabled = False

If DisplayType = "LABEL" Then
    picSystem.Left = picLabel.Left
    picSystem.Top = picLabel.Top
Else
    picLabel.Left = picSystem.Left
    picLabel.Top = picSystem.Top
End If

End Sub

Private Sub picHeart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu mnuMenu
End If

End Sub

Private Sub picLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Pos = "LEFT"

Timer2.Enabled = True

txtDefine.Text = ""

End Sub

Private Sub picLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer2.Enabled = False

End Sub

Private Sub picPlay_Click()

Tick = False
Timer3.Enabled = True

picPlay.Visible = False
picStop.Visible = True

txtDefine.Text = ""

End Sub

Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Pos = "RIGHT"
Timer2.Enabled = True

txtDefine.Text = ""

End Sub

Private Sub picRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer2.Enabled = False

End Sub

Private Sub picShowLabel_Click()

DisplayType = "LABEL"
picLabel.Visible = True
picSystem.Visible = False

picShowLabel.Visible = False
picShowSystem.Visible = True

End Sub

Private Sub picShowSystem_Click()

DisplayType = "SYSTEM"
picLabel.Visible = False
picSystem.Visible = True

picShowLabel.Visible = True
picShowSystem.Visible = False

End Sub


Private Sub picStop_Click()

Timer3.Enabled = False

picPlay.Visible = True
picStop.Visible = False

End Sub

Private Sub rtf_Change()

txtDefine.Text = rtf.Text

End Sub

Private Sub Timer1_Timer()

Select Case ButtonClick
    
    Case Is = "UP"
        If DisplayType = "LABEL" Then
            If picLabel.Top < 0 Then
                picLabel.Top = picLabel.Top + 120
            End If
        Else
            If picSystem.Top < 0 Then
                picSystem.Top = picSystem.Top + 120
            End If
        End If

    Case Is = "DOWN"
        If DisplayType = "LABEL" Then
            If picLabel.Top > -11040 Then
                picLabel.Top = picLabel.Top - 120
            End If
        Else
            If picSystem.Top > -11040 Then
                picSystem.Top = picSystem.Top - 120
            End If
        End If
    
    Case Is = "LEFT"
        If DisplayType = "LABEL" Then
            If picLabel.Left < 0 Then
                picLabel.Left = picLabel.Left + 120
            End If
        Else
            If picSystem.Left < 0 Then
                picSystem.Left = picSystem.Left + 120
            End If
        End If

    Case Is = "RIGHT"
        If DisplayType = "LABEL" Then
            If picLabel.Left > -2040 Then
                picLabel.Left = picLabel.Left - 120
            End If
        Else
            If picSystem.Left > -2040 Then
                picSystem.Left = picSystem.Left - 120
            End If
        End If

End Select

End Sub

Private Sub Timer2_Timer()

If Pos = "LEFT" Then
    
    If Beat = 1 Then
        Beat = 10
    Else
        Beat = Beat - 1
    End If

Else
    
    If Beat = 10 Then
        Beat = 1
    Else
        Beat = Beat + 1
    End If

End If

picHeart.Picture = LoadPicture(App.Path & _
        "\pictures\heart\" & Trim(Str(Beat)) & ".1.jpg")

End Sub

Private Sub Timer3_Timer()

If Tick = True Then
    picHeart.Picture = LoadPicture(App.Path & _
        "\pictures\heart\" & LTrim(Str(Beat)) & ".2.jpg")
    Tick = False
Else
    picHeart.Picture = LoadPicture(App.Path & _
        "\pictures\heart\" & LTrim(Str(Beat)) & ".1.jpg")
    Tick = True
End If

End Sub
