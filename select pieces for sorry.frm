VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Choose pawns"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "Test"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Blue"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Tag             =   "1"
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yellow"
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Tag             =   "16"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Red"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Tag             =   "46"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Green"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Tag             =   "31"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image sorrypieces 
      Height          =   360
      Index           =   0
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Person's Name"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'For x = 0 To 3
On Error Resume Next

Load Form1


'Form1.Show

End Sub

Private Sub Form_Load()
For x = 1 To 4
'MsgBox x

Load sorrypieces(x)
sorrypieces(x).Picture = LoadPicture(App.Path & "\" & Check1(x - 1).Caption & " game piece.jpg")
sorrypieces(x).Left = Check1(x - 1).Left - sorrypieces(x).Width - 50
sorrypieces(x).Top = Check1(x - 1).Top
sorrypieces(x).Visible = True
Next
Me.Show

End Sub
