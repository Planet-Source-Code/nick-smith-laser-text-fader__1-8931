VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Text Fader with Laser Writer"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H006B6529&
      ForeColor       =   &H00C0C0C0&
      Height          =   3960
      ItemData        =   "txtFade.frx":0000
      Left            =   4245
      List            =   "txtFade.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Draw Mode"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "Laser Write"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   2025
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ReDraw"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "Faded Text"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counter2 As Long

Private Sub Command1_Click()
Picture1.Cls
Picture4.Cls
Picture1.FontName = List1.List(List1.ListIndex)
TextOut Picture1.hDC, 0, Picture1.ScaleHeight / 2 - 20, Text1.Text, Len(Text1.Text)

tmpRed1& = Get_RGB(Picture2.BackColor, 1)
tmpGreen1& = Get_RGB(Picture2.BackColor, 2)
tmpBlue1& = Get_RGB(Picture2.BackColor, 3)
tmpRed2& = Get_RGB(Picture3.BackColor, 1)
tmpGreen2& = Get_RGB(Picture3.BackColor, 2)
tmpBlue2& = Get_RGB(Picture3.BackColor, 3)
For Counter1 = Picture1.ScaleWidth To 0 Step -1
    For Counter2 = (Picture1.ScaleHeight / 2) - 20 To (Picture1.ScaleHeight / 2) + 20
    If GetPixel(Picture1.hDC, Counter1, Counter2) = Picture1.ForeColor Then
    tmpDir% = Counter1
    tmpLeave% = 1
    Exit For
    End If
    Next Counter2
If tmpLeave% = 1 Then Exit For
Next Counter1

rd = (tmpRed2& - tmpRed1&) / tmpDir%
gr = (tmpGreen2& - tmpGreen1&) / tmpDir%
bl = (tmpBlue2& - tmpBlue1&) / tmpDir%

For Counter1 = 0 To Picture1.ScaleWidth
    For Counter2 = 0 To Picture1.ScaleHeight
    If GetPixel(Picture1.hDC, Counter1, Counter2) = Picture1.ForeColor Then
    finalRed& = tmpRed1& + (rd * Counter1)
    finalGreen& = tmpGreen1& + (gr * Counter1)
    finalBlue& = tmpBlue1& + (bl * Counter1)
    
    finalCol! = RGB(finalRed&, finalGreen&, finalBlue&)
        If Option2.Value = True Then
        Picture4.Line (Picture4.ScaleWidth, Picture4.ScaleHeight)-(Counter1, Counter2), finalCol!
        SetPixel Picture4.hDC, Counter1, Counter2, finalCol!
        hold (0.0001)
        Picture4.Line (Picture4.ScaleWidth, Picture4.ScaleHeight)-(Counter1, Counter2), Picture4.BackColor
        Else
        SetPixel Picture4.hDC, Counter1, Counter2, finalCol!
        End If
    End If
    Next Counter2
Next Counter1
End Sub

Private Sub Form_Load()
Option2.Value = True
For Counter1 = 0 To Screen.FontCount - 1
List1.AddItem Screen.Fonts(Counter1)
Next Counter1
List1.ListIndex = 0
End Sub


Private Sub Picture2_Click()
CommonDialog1.ShowColor
Picture2.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture3_Click()
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
End Sub
