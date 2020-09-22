VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Sorry"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   6600
      TabIndex        =   37
      Top             =   3720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton rollback 
      Caption         =   "Roll Back Turn"
      Height          =   375
      Left            =   6720
      TabIndex        =   36
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Skip Turn"
      Height          =   375
      Left            =   6720
      TabIndex        =   35
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Debug"
      Height          =   255
      Left            =   6720
      TabIndex        =   34
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton redos 
      Caption         =   "Redo Turn"
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Back 1"
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Forward 10"
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Text            =   "0"
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Possible Moves"
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton pass 
      Caption         =   "Pass"
      Height          =   255
      Left            =   9480
      TabIndex        =   24
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   8280
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Must have 1 or 2 or someone to sorry to go"
      Height          =   255
      Left            =   9960
      TabIndex        =   16
      Top             =   7920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Current Status"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Current Safety Position"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Current Board Position"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Messages"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Trade Places"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Text            =   "7"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   9480
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Current Turn"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Sorry"
      Height          =   255
      Left            =   10200
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   9960
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Movement"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Index           =   2
      Left            =   0
      Picture         =   "sorry.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Image sorrypieces 
         Height          =   360
         Index           =   0
         Left            =   4440
         Top             =   3480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Image5 
         Height          =   1290
         Left            =   3480
         ToolTipText     =   "Move Forward One Space Or Take A Man Out Of Start"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   1290
         Index           =   2
         Left            =   2520
         Picture         =   "sorry.frx":8069
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.Label todolabel 
      Alignment       =   2  'Center
      Caption         =   "To Do:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label status 
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Additions"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label todo 
      Height          =   855
      Left            =   6720
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label color 
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label player 
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Turn:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Choose Next Card"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Number for first pawn for splits"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   10680
      TabIndex        =   1
      Top             =   8160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As FileSystemObject
Private t
Private card() As cardinfo
Private nextcard
Private gamepiece() As pieceinfo
Private alreadyasked As Boolean
Private turnnum As Integer
Private players() As playerinfo
Private slidenum As Integer
Private manuelturn As manuelinfo
Private willmoveout As moveoutinfo
Private testing As Boolean
Private rollbacks As Boolean
Private unloadthis As Boolean
Private previousturn() As pieceinfo


Private secondss As Integer


Private Type manuelinfo
firstpawn As Integer
secondpawn As Integer
value As Integer
samecolor As Integer
opponent As Integer
card As Integer
End Type


Private Type possiblemove
pawn As Integer
playnum As Integer
safety As Boolean
leftss As Integer
End Type
Private Type moveoutinfo
willmove As Boolean
pawn As Integer
End Type



Private Type moveinfo
optional As Integer
required As Integer
firstpawn As Integer
secondpawn As Integer
value As Integer

End Type


Private Type playerinfo
color As String
playername As String
startnum As Integer
End Type


Private Type positioning
leftss As Integer
topss As Integer
noprocess As Boolean
End Type

Private previous() As pieceinfo
Private opponent As pieceinfo
Private continues As Boolean

Private Type splitinformation
cansplit As Boolean
firstpawn As Integer
secondpawn As Integer
firstposition As Integer
secondposition As Integer
End Type



Private y As Integer


Private Type pieceinfo
color As String
status As String
starting As Integer
currentposition As Integer
boardposition As Integer
safetyposition As Integer
leftss As Integer
topss As Integer
gamepiece As Integer

End Type


Private Type cardinfo
number As Integer
card As Integer
'optional As Boolean
description As String


'forward As Boolean
'sorry As Boolean
'trade As Boolean
'split As Boolean
'backward As Boolean
End Type

Sub shufflecards()
'ReDim cards(52)
ReDim card(45)
'Open "C:\E.txt" For Output As #1
Dim y As Integer


For x = 1 To 45
'MsgBox tempcard.NumberOfCards
y = 0

Do

Randomize
ask1 = Int((45 * Rnd) + 1)

'If ask1 = 1 Then
'MsgBox "check this"
'End If

'cards(x) = ask1


dets = repeats(ask1, x)
'MsgBox dets

If dets = "False" Then
'card(x).number = ask1
'MsgBox x
'MsgBox ask1
'Write #1, ask1
y = y + 1
If y = 4 Then


card(x).number = ask1

Exit Do
End If
End If

Loop

Next

'Close #1






End Sub
'Sub testings()
'MsgBox tempcard.PointValue(15)

'End Sub

Function repeats(ask1, numbers)
For x = 1 To numbers
If ask1 = card(x).number Then
'MsgBox ask1
'MsgBox cards(x)
'MsgBox x

repeats = "True"
Exit Function
Exit For
End If
Next
repeats = "False"


End Function

Private Function sorrycard(cards As Integer) As cardinfo
'For X = 1 To 45
sorrycard.number = cards
newx = 5
If cards <= newx Then
sorrycard.card = 1
sorrycard.description = "Move from Start or move forward 1."
'sorrycard.sorry = False
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 2
sorrycard.description = "Move from the Start or move forward 2. Draw Again."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 3
sorrycard.description = "Move forward 3."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 4
sorrycard.description = "Move backward 4."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 5
sorrycard.description = "Move forward 5."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 7
sorrycard.description = "Move forward 7 or split between two pawns."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 8
sorrycard.description = "Move forward 8."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 10
sorrycard.description = "Move forward 10 or move backward 1."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 11
sorrycard.description = "Move forward 11 or change places with an opponent."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 12
sorrycard.description = "Move forward 12."
Exit Function
End If
newx = newx + 4
If cards <= newx Then
sorrycard.card = 13
sorrycard.description = "Move from Start and switch places with an opponent, who you bump back to Start."
Exit Function
End If



End Function
Private Function tradeplaces(ngamepiece As Integer, getpiece As Integer) As positioning
'MsgBox "Test"

position1 = gamepiece(ngamepiece).boardposition
topss1 = sorrypieces(ngamepiece).Top
leftss1 = sorrypieces(ngamepiece).Left

position2 = gamepiece(getpiece).boardposition
topss2 = sorrypieces(getpiece).Top
leftss2 = sorrypieces(getpiece).Left
opponent.boardposition = gamepiece(getpiece).boardposition
opponent.safetyposition = gamepiece(getpiece).safetyposition
opponent.status = gamepiece(getpiece).status
opponent.color = gamepiece(getpiece).color
opponent.gamepiece = getpiece
opponent.currentposition = gamepiece(getpiece).currentposition
opponent.starting = gamepiece(getpiece).starting
opponent.leftss = sorrypieces(getpiece).Left
gamepiece(ngamepiece).boardposition = position2
gamepiece(getpiece).boardposition = position1
'newx = gamepiece(ngamepiece).boardposition - 1
searches = InStr(List1.List(newx), " ")
tradeplaces.leftss = leftss2
tradeplaces.topss = topss2
gamepiece(ngamepiece).currentposition = currentposition(ngamepiece, getpiece)
gamepiece(getpiece).currentposition = currentposition(getpiece, ngamepiece)



'MsgBox opponent.boardposition & " for board position of opponent"
'MsgBox opponent.currentposition & " for current position of opponent"

If sliderule(ngamepiece, True, True) = True Then
tradeplaces.noprocess = True
End If
'MsgBox gamepiece(ngamepiece).boardposition


If continues = True Then

newx = gamepiece(ngamepiece).boardposition - 1
searches = InStr(List1.List(newx), " ")
'leftss = topss1
'topss = leftss1
If sliderule(getpiece, False, True) = False Then

sorrypieces(getpiece).Left = leftss1
sorrypieces(getpiece).Top = topss1
End If
End If



'MsgBox gamepiece(ngamepiece).currentposition
'MsgBox gamepiece(getpiece).currentposition




End Function
Private Function currentposition(ngamepiece, getpiece) As Integer
For x = 0 To 59
z = x + gamepiece(ngamepiece).starting
If z > 59 Then
z = z - 60
End If


searches = InStr(List1.List(z), " ")
leftss = Mid(List1.List(z), 1, searches - 1)
topss = Mid(List1.List(z), searches, Len(List1.List(z)) - searches + 1)
'MsgBox sorrypieces(getpiece).Left & "   " & leftss & vbCrLf & sorrypieces(getpiece).Top & "   " & topss & vbCrLf & z

'MsgBox z

'Write #1, z



If sorrypieces(getpiece).Left = leftss And sorrypieces(getpiece).Top = topss Then
currentposition = x + 1

'gamepiece(ngamepiece).currentposition = x + 1
'gamepiece(ngamepiece).status = "Out"
'MsgBox gamepiece(ngamepiece).currentposition
'MsgBox z
'MsgBox gamepiece(ngamepiece).currentposition & "   " & x & "   " & gamepiece(ngamepiece).starting
If Check4.value = 1 Then

MsgBox currentposition & " so far"
End If

If currentposition = 60 Then
currentposition = 0
End If



'MsgBox z
'MsgBox leftss
'MsgBox topss

'backtostart (getpiece)
Exit For
End If
Next
End Function
Private Function sliderule(ngamepiece As Integer, isoptional As Boolean, doprocess As Boolean) As Boolean


Dim gameposition As positioning

'gameposition = newposition(ngamepiece, z)

'MsgBox gameposition.topss & "  " & gameposition.leftss
'MsgBox gamepiece(ngamepiece).boardposition

If gamepiece(ngamepiece).boardposition = 58 And gamepiece(ngamepiece).currentposition <> gamepiece(ngamepiece).boardposition And gamepiece(ngamepiece).color <> "Blue" Then
'MsgBox "Test"
If doprocess = True Then

gameposition = sliding(ngamepiece, True, 3, isoptional)
Else
slidenum = 3
End If


'MsgBox gamepiece(ngamepiece).boardposition
'MsgBox gamepiece(ngamepiece).currentposition
sliderule = True

ElseIf gamepiece(ngamepiece).currentposition = 12 Or gamepiece(ngamepiece).currentposition = 27 Or gamepiece(ngamepiece).currentposition = 42 Then
'MsgBox "Test"
If doprocess = True Then

gameposition = sliding(ngamepiece, False, 3, isoptional)
Else
slidenum = 3
End If

sliderule = True

ElseIf gamepiece(ngamepiece).currentposition = 20 Or gamepiece(ngamepiece).currentposition = 35 Or gamepiece(ngamepiece).currentposition = 50 Then
'MsgBox "Test"
If doprocess = True Then
gameposition = sliding(ngamepiece, False, 4, isoptional)
Else
slidenum = 4
End If

sliderule = True

End If
'MsgBox "Test so far"
'If Check4.value = 1 Then
'MsgBox gamepiece(ngamepiece).currentposition & " for current position"
'End If

If sliderule = True And doprocess = True Then
sorrypieces(ngamepiece).Top = gameposition.topss
sorrypieces(ngamepiece).Left = gameposition.leftss
End If


End Function
Private Function splitinformation(ngamepiece As Integer, firstpos, secondpiece As Integer) As splitinformation

Dim secondpieces As Integer
Dim homenumbers As Integer

secondpieces = 0

For z = 1 To UBound(gamepiece)
If z <> ngamepiece And gamepiece(z).color = gamepiece(ngamepiece).color And gamepiece(z).status <> "Home" And gamepiece(z).status <> "Start" Then
secondpieces = z
Exit For
End If
Next
If secondpiece <> 0 Then
secondpieces = secondpiece
End If

If secondpieces = 0 Then
splitinformation.cansplit = False
splitinformation.firstpawn = ngamepiece
splitinformation.firstposition = 7
Exit Function
End If
If gamepiece(ngamepiece).status = "Safety" And gamepiece(secondpieces).status = "Safety" Then
homenumbers = calculatehomes(ngamepiece, secondpiece)
'MsgBox homenumbers

If homenumbers < 7 Then
splitinformation.cansplit = False
splitinformation.firstpawn = ngamepiece
splitinformation.firstposition = 7
Exit Function
End If
r = 0

For s = 5 To 1 Step -1
r = r + 1
'MsgBox x
'MsgBox y

If gamepiece(ngamepiece).safetyposition = s Then
home1 = r
Exit For
End If
Next
'MsgBox home1 & "   " & firstpos

If CInt(home1) < CInt(firstpos) Then
splitinformation.cansplit = False
splitinformation.firstpawn = ngamepiece
splitinformation.firstposition = 7
Exit Function
End If
End If

splitinformation.cansplit = True
splitinformation.firstpawn = ngamepiece
splitinformation.firstposition = firstpos
splitinformation.secondpawn = secondpieces
splitinformation.secondposition = 7 - firstpos


If Check4.value = 1 Then
MsgBox splitinformation.firstpawn & "  for first pawn  " & splitinformation.firstposition & "   for first position  " & splitinformation.secondpawn & "   for second pawn  " & splitinformation.secondposition & "   for second position"
End If





Exit Function








'If z <> y + 1 And gamepiece(z).color <> gamepiece(y).color And gamepiece(z).status = "Out" Then



End Function
Private Function splitmove(splitinfo As splitinformation) As positioning()
Dim newsplit(2) As positioning
Dim newx As Integer
Dim ngamepiece As Integer
Dim newinfo As Integer
Dim exits As Boolean
exits = False
If Check4.value = 1 Then
MsgBox splitinfo.firstposition & " for first position"
End If

For xx = 1 To 2
'If xx = 2 Then MsgBox "Test"
If xx = 2 And rollbacks = True Then
'splitmove.leftss = previous(2).leftss
'splitmove.topss = previous(2).topss
newsplit(2).leftss = previous(2).leftss
newsplit(2).topss = previous(2).topss
splitmove = newsplit

Exit Function
Exit For
End If

If xx = 1 Then

newx = splitinfo.firstposition
ngamepiece = splitinfo.firstpawn
Else
newx = splitinfo.secondposition
ngamepiece = splitinfo.secondpawn
End If
Dim oldpositions As Integer

oldpositions = gamepiece(ngamepiece).boardposition
If gamepiece(ngamepiece).status = "Safety" Then
newinfo = newx

'newsplit(xx) = safeties(ngamepiece, newinfo, True, oldpositions)
newsplit(xx) = safeties(ngamepiece, newx, False, 0)
Else



If gamepiece(ngamepiece).status = "Out" And exits = False Then


gamepiece(ngamepiece).boardposition = gamepiece(ngamepiece).boardposition + newx
End If







'MsgBox "Test"





'MsgBox newx

'possible fix
newx = gamepiece(ngamepiece).boardposition
'if not then take out above code


If newx > List1.ListCount Then
newx = newx - List1.ListCount
If gamepiece(ngamepiece).status = "Out" And exits = False Then


gamepiece(ngamepiece).boardposition = newx
End If

End If
If newx < 1 Then
newx = 60 + newx



End If
'MsgBox gamepiece(ngamepiece).currentposition
newcard = newx
'possible fix
If xx = 1 Then
newcard = splitinfo.firstposition
Else
newcard = splitinfo.secondposition
End If




If gamepiece(ngamepiece).currentposition + newcard > 58 And gamepiece(ngamepiece).currentposition <> 59 And exits = False Then

If Check4.value = 1 Then
MsgBox newcard & " for new card for splits " & gamepiece(ngamepiece).currentposition & " for position"

End If

'MsgBox gamepiece(ngamepiece).currentposition + newx



'newinfo = gamepiece(ngamepiece).currentposition + newcard
newinfo = newcard - 1





newinfo = 58 - gamepiece(ngamepiece).currentposition
newinfo = newcard - newinfo
newinfo = newinfo - 1

If Check4.value = 1 Then
MsgBox newinfo & " for new information"

End If



'newposition = safeties(ngamepiece, newinfo, True, oldpositions)
newsplit(xx) = safeties(ngamepiece, newinfo, True, oldpositions)


'MsgBox "You will be safe"
'gamepiece(ngamepiece).boardposition = oldpositions

'newposition.leftss = sorrypieces(ngamepiece).Left
'newposition.topss = sorrypieces(ngamepiece).Top
'Exit Function
'had end if

ElseIf gamepiece(ngamepiece).status = "Start" Then
gamepiece(ngamepiece).currentposition = 0
newsplit(xx).leftss = sorrypieces(ngamepiece).Left
newsplit(xx).topss = sorrypieces(ngamepiece).Top

Else

'MsgBox currentposition

If gamepiece(ngamepiece).currentposition = 59 And exits = False Then
gamepiece(ngamepiece).currentposition = card(nextcard - 1).card - 1
ElseIf exits = True Then
gamepiece(ngamepiece).currentposition = gamepiece(ngamepiece).currentposition

Else



gamepiece(ngamepiece).currentposition = gamepiece(ngamepiece).currentposition + newcard
If gamepiece(ngamepiece).currentposition < 1 And gamepiece(ngamepiece).status <> "Start" And exits = False Then

gamepiece(ngamepiece).currentposition = 60 + gamepiece(ngamepiece).currentposition
End If

End If

'If gamepiece(ngamepiece).currentposition > 58 Then
'MsgBox card(nextcard - 1).card

'gamepiece(ngamepiece).currentposition = card(nextcard - 1).card
'End If
'MsgBox newx
'MsgBox gamepiece(ngamepiece).currentposition

newx = gamepiece(ngamepiece).boardposition

newx = newx - 1
'MsgBox newx

searches = InStr(List1.List(newx), " ")


If continues = False Then

newsplit(xx).leftss = previous(xx).leftss
newsplit(xx).topss = previous(xx).topss

ElseIf gamepiece(ngamepiece).status = "Start" Then
'MsgBox ngamepiece

newsplit(xx).leftss = sorrypieces(ngamepiece).Left
newsplit(xx).topss = sorrypieces(ngamepiece).Top

'gamepiece(newpiece).boardposition = previous(xx).boardposition
'gamepiece(newpiece).currentposition = previous(xx).currentposition
'gamepiece(newpiece).status = previous(xx).status
'gamepiece(newpiece).safetyposition = previous(xx).safetyposition

ElseIf acceptablemove(ngamepiece, False) = True Then

'newsplit(xx).leftss = Mid(List1.List(newx), 1, searches - 1)



'newsplit(xx).leftss = Mid(List1.List(newx), 1, searches - 1)
'MsgBox newx
newsplit(xx).leftss = Mid(List1.List(newx), 1, searches - 1)

newsplit(xx).topss = Mid(List1.List(newx), searches, Len(List1.List(newx)) - searches + 1)



'newposition.leftss = Mid(List1.List(newx), 1, searches - 1)
'newposition.topss = Mid(List1.List(newx), searches, Len(List1.List(newx)) - searches + 1)


Else



newsplit(xx).leftss = previous(xx).leftss
newsplit(xx).topss = previous(xx).topss

If xx = 1 Then
exits = True
End If



'newposition.leftss = previous(1).leftss
'newposition.topss = previous(1).topss
End If
End If
End If
'MsgBox "split " & xx

Next
splitmove = newsplit



'MsgBox gamepiece(ngamepiece).currentposition









End Function

Private Function calculatehomes(ngamepiece As Integer, newpiece As Integer) As Integer
r = 0
'MsgBox "Test 2"

For s = 5 To 1 Step -1
r = r + 1
'MsgBox x
'MsgBox y
'MsgBox gamepiece(ngamepiece).safetyposition & "   pos " & s & "   pos"


If gamepiece(ngamepiece).safetyposition = s Then
home1 = r
Exit For
End If
Next

r = 0

For s = 5 To 1 Step -1
r = r + 1
'MsgBox x
'MsgBox y

If gamepiece(newpiece).safetyposition = s Then
home2 = r
Exit For
End If
Next
'MsgBox home1
'MsgBox home2

calculatehomes = home1 + home2


End Function
Private Sub Command1_Click()
alreadyasked = False

continues = True

On Error Resume Next
Close #1
On Error GoTo 0
If IsNumeric(Text1.Text) = False Then
Text1.Text = 7
End If

Dim gameposition As positioning
Dim numbers As Integer

'MsgBox card(nextcard - 1).card & vbCrLf & card(nextcard - 1).description
'For z = 0 To List2.ListCount - 1
'If List2.Selected(z) = True Then
'y = z + 1
'Exit For
'End If
'Next
y = players(turnnum).startnum + Text3.Text

If Check3.value = 1 And card(nextcard - 1).card <> 1 And card(nextcard - 1).card <> 2 And card(nextcard - 1).card <> 13 And gamepiece(y).status = "Start" Then
MsgBox "Sorry, you must have a 1 or 2 to start or use the sorry card on someone"
Exit Sub
End If


numbers = 0
ReDim previous(1)
previous(1).gamepiece = y
previous(1).boardposition = gamepiece(y).boardposition
previous(1).safetyposition = gamepiece(y).safetyposition
previous(1).status = gamepiece(y).status
previous(1).color = gamepiece(y).color
previous(1).currentposition = gamepiece(y).currentposition
previous(1).starting = gamepiece(y).starting
previous(1).leftss = sorrypieces(y).Left
previous(1).topss = sorrypieces(y).Top

If card(nextcard - 1).card = 7 And Text1.Text < 7 And gamepiece(y).status <> "Home" And gamepiece(y).status <> "Start" Then
'MsgBox "split so far"

'new information
Dim splits As splitinformation
'MsgBox "Test"

splits = splitinformation(y, Text1.Text)

If splits.cansplit = False Then
MsgBox "Sorry, you cannnot use this split"
Exit Sub
End If

'MsgBox splits.firstpawn & "   " & splits.firstposition
'MsgBox splits.secondpawn & "   " & splits.secondposition



ReDim Preserve previous(2)
previous(2).gamepiece = splits.secondpawn
previous(2).boardposition = gamepiece(splits.secondpawn).boardposition
previous(2).currentposition = gamepiece(splits.secondpawn).currentposition
previous(2).leftss = sorrypieces(splits.secondpawn).Left
previous(2).topss = sorrypieces(splits.secondpawn).Top
previous(2).status = gamepiece(splits.secondpawn).status


Dim thissplit() As positioning
thissplit = splitmove(splits)
For xx = 1 To 2
'MsgBox thissplit(xx).leftss & " for left " & xx & "  " & thissplit(xx).topss & " for top " & xx


If xx = 1 Then
sorrypieces(splits.firstpawn).Left = thissplit(1).leftss
sorrypieces(splits.firstpawn).Top = thissplit(1).topss
Else
sorrypieces(splits.secondpawn).Left = thissplit(2).leftss
sorrypieces(splits.secondpawn).Top = thissplit(2).topss
End If
Next




If sliderule(splits.firstpawn, True, True) = False And continues = True Then
sorrypieces(ngamepiece).Left = thissplit(1).leftss
sorrypieces(ngamepiece).Top = thissplit(1).topss
End If
If continues = False Then
sorrypieces(splits.secondpawn).Left = previous(2).leftss
sorrypieces(splits.secondpawn).Top = previous(2).topss


Exit Sub
End If

If continues = True And sliderule(splits.secondpawn, True, True) = False Then
sorrypieces(splits.secondpawn).Left = thissplit(2).leftss
sorrypieces(splits.secondpawn).Top = thissplit(2).topss
End If


Exit Sub




'put into array




End If


If card(nextcard - 1).card = 11 And Check1.value = 1 Then
For z = 1 To UBound(gamepiece)
If z <> y + 1 And gamepiece(z).color <> gamepiece(y).color And gamepiece(z).status = "Out" Then
numbers = z
Exit For
End If
Next
If numbers = 0 Then
MsgBox "No one to trade places"
Exit Sub
End If
opponent.boardposition = gamepiece(numbers).boardposition
opponent.safetyposition = gamepiece(numbers).safetyposition
opponent.status = gamepiece(numbers).status
opponent.color = gamepiece(numbers).color
opponent.gamepiece = y
opponent.currentposition = gamepiece(numbers).currentposition
opponent.starting = gamepiece(numbers).starting
opponent.leftss = sorrypieces(numbers).Left

opponent.topss = sorrypieces(numbers).Top

gameposition = tradeplaces(y, numbers)

If gameposition.noprocess = False Then
sorrypieces(y).Top = gameposition.topss
sorrypieces(y).Left = gameposition.leftss
End If


Exit Sub
End If



If card(nextcard - 1).card = 13 Then
If card(nextcard - 1).card = 13 And gamepiece(y).status <> "Start" Then
MsgBox "Sorry, you must use a man from start for the sorry"
Exit Sub
End If

For z = 1 To UBound(gamepiece)
If z <> y + 1 And gamepiece(z).color <> gamepiece(y).color And gamepiece(z).status = "Out" Then
numbers = z
Exit For
End If
Next

If numbers = 0 Then
MsgBox "No one to sorry"
Exit Sub
End If
End If


'y = 5
'MsgBox gamepiece(y).boardposition

z = numbers



If gamepiece(y).status <> "Home" Then
'MsgBox z

opponent.boardposition = gamepiece(z).boardposition
opponent.safetyposition = gamepiece(z).safetyposition
opponent.status = gamepiece(z).status
opponent.color = gamepiece(z).color
opponent.currentposition = gamepiece(z).currentposition
opponent.starting = gamepiece(z).starting
opponent.leftss = sorrypieces(z).Left
opponent.topss = sorrypieces(z).Top
opponent.gamepiece = z

gameposition = newposition(y, z, 0)


If sliderule(y, True, True) = False Then


'MsgBox gamepiece(y).currentposition






sorrypieces(y).Top = gameposition.topss
sorrypieces(y).Left = gameposition.leftss
'MsgBox gamepiece(y).boardposition
End If

Else
MsgBox "Sorry, you are already home"
End If

End Sub

Private Sub Command2_Click()
Dim gameposition As positioning
'y = 1
'MsgBox "Test so far"





gameposition = newposition(y, 1)
'MsgBox gameposition.topss & "  " & gameposition.leftss

sorrypieces(y).Top = gameposition.topss
sorrypieces(y).Left = gameposition.leftss
'MsgBox gamepiece(y).currentposition

End Sub

Private Sub Command3_Click()
'For z = 0 To List2.ListCount - 1
'If List2.Selected(z) = True Then
'y = z + 1
'Exit For
'End If
'Next

'MsgBox gamepiece(y).currentposition
'MsgBox turnnum

End Sub

Private Sub Command4_Click()
For z = 0 To List2.ListCount - 1
If List2.Selected(z) = True Then
y = z + 1
Exit For
End If
Next

MsgBox gamepiece(y).boardposition

End Sub

Private Sub Command5_Click()
For z = 0 To List2.ListCount - 1
If List2.Selected(z) = True Then
y = z + 1
Exit For
End If
Next

MsgBox gamepiece(y).safetyposition

End Sub

Private Sub Command6_Click()
For z = 0 To List2.ListCount - 1
If List2.Selected(z) = True Then
y = z + 1
Exit For
End If
Next

MsgBox gamepiece(y).status



End Sub


Private Function startings(x) As Integer
z = 0
'MsgBox "Check so far"
Dim newx As Integer
newx = x

Do
'MsgBox z
'MsgBox gamepiece(z).status
'w = z + x
'MsgBox w
z = z + 1
'MsgBox z


'MsgBox z
'MsgBox gamepiece(z).color & "   " & gamepiece(z).status


'If gamepiece(x).status <> "Start" Then
'if sorrypieces(x).Left=

'z = z + 1
startings = z

'MsgBox z
Exit Function
Exit Do
End If
newx = newx + 1
Loop

End Function
Private Function occupies(x, leftss, topss) As Boolean
occupies = False
For z = x To x + 3
'MsgBox z
'MsgBox x + 4
If sorrypieces(z).Top = topss And sorrypieces(z).Left = leftss Then
occupies = True
Exit Function
Exit For
End If
Next

End Function

Sub backtostart(ngamepiece As Integer)
'MsgBox ngamepiece

'MsgBox "Test so far"
Dim z As Integer

For x = 1 To UBound(gamepiece)
If gamepiece(x).color = gamepiece(ngamepiece).color Then
'MsgBox x

'z = startings(x)
Exit For
End If


Next


On Error Resume Next
Close #1
On Error GoTo 0

Open App.Path & "\" & gamepiece(ngamepiece).color & " start.txt" For Input As #1
'MsgBox z

'MsgBox z

Do

'For x = 1 To z
Line Input #1, c
'Next
searches = InStr(c, "   ")
leftss = Mid(c, 1, searches - 1)
topss = Mid(c, searches, Len(c) - searches + 1)
If occupies(x, leftss, topss) = False Then
Exit Do
End If
Loop

Close #1


'Load sorrypieces(y)
'sorrypieces(y).Picture = LoadPicture(app.path & "\" & Form2.Check1(x).Caption & " game piece.jpg")

gamepiece(ngamepiece).boardposition = 0
gamepiece(ngamepiece).currentposition = 0
If gamepiece(ngamepiece).status = "Safety" Then
MsgBox "You should check this because it could be another hidden bug"
End If

If gamepiece(ngamepiece).status = "Safety" Then
MsgBox "You should check this because it could be another hidden bug"
MsgBox ngamepiece & " game piece is being sent back to start for unexplained reason"

End If

gamepiece(ngamepiece).status = "Start"


'Set sorrypieces(y).Container = Picture1(2)
sorrypieces(ngamepiece).Top = topss
sorrypieces(ngamepiece).Left = leftss
'MsgBox ngamepiece


End Sub
Private Function safeties(ngamepiece As Integer, starts As Integer, remaining As Boolean, oldpositions As Integer) As positioning
'MsgBox starts
'MsgBox gamepiece(ngamepiece).currentposition



Dim x As Integer
If remaining = True And starts = 0 Then
x = starts + 1
ElseIf remaining = False And starts <> 0 Then
x = gamepiece(ngamepiece).safetyposition + starts

ElseIf card(nextcard - 1).card = 4 Then
x = -4 + gamepiece(ngamepiece).safetyposition
ElseIf remaining = True And starts <> 0 Then
x = starts + 1
ElseIf card(nextcard - 1).card = 10 Then
x = -1 + gamepiece(ngamepiece).safetyposition
Else
x = card(nextcard - 1).card + gamepiece(ngamepiece).safetyposition

End If
If Check4.value = 1 Then
MsgBox x & " for starts"
End If

'MsgBox x


If x < 1 Then

'MsgBox "No longer safe   " & x


'later fix

'If remaining = True Then

'gamepiece(ngamepiece).boardposition = oldpositions
'safeties.leftss = sorrypieces(ngamepiece).Left
'safeties.topss = sorrypieces(ngamepiece).Top
'Exit Function
'MsgBox x

'Else
'MsgBox x

newxx = gamepiece(ngamepiece).starting + 58
'MsgBox newxx & " 1" & "   " & List1.ListCount


newxx = newxx + x
'MsgBox newxx

If newxx > List1.ListCount Then
newxx = newxx - List1.ListCount



End If
'MsgBox newxx & " 2"


If 58 - x = 59 Then
currentss = 57
'newxx = newxx - 1
Else
currentss = 58 + x

'newxx = newxx
End If
'MsgBox newxx & " 3"




gamepiece(ngamepiece).boardposition = newxx
gamepiece(ngamepiece).currentposition = currentss
gamepiece(ngamepiece).status = "Out"
gamepiece(ngamepiece).safetyposition = 0



newx = newxx - 1
'MsgBox newx

searches = InStr(List1.List(newx), " ")

If acceptablemove(ngamepiece, False) = True Then



safeties.leftss = Mid(List1.List(newx), 1, searches - 1)
safeties.topss = Mid(List1.List(newx), searches, Len(List1.List(newx)) - searches + 1)
Else
safeties.leftss = previous(1).leftss
safeties.topss = previous(1).topss
End If


'MsgBox newxx
'End If

'safeties.leftss = sorrypieces(ngamepiece).Left
'safeties.topss = sorrypieces(ngamepiece).Top
'end fix
Exit Function
End If

If x = 6 Then


'MsgBox "Test home"



For x = 1 To UBound(gamepiece)
If gamepiece(x).color = gamepiece(ngamepiece).color Then
'MsgBox x

'z = startings(x)
Exit For
End If


Next


On Error Resume Next
Close #1
On Error GoTo 0

Open App.Path & "\" & gamepiece(ngamepiece).color & " home.txt" For Input As #1
'MsgBox z

'MsgBox z

Do

'For x = 1 To z
Line Input #1, c
'Next
searches = InStr(c, "   ")
leftss = Mid(c, 1, searches - 1)
topss = Mid(c, searches, Len(c) - searches + 1)
If occupies(x, leftss, topss) = False Then
Exit Do
End If
Loop

Close #1







'For q = 1 To UBound(gamepiece)
'If gamepiece(q).color = gamepiece(ngamepiece).color Then
'z = 0
'Do
'MsgBox z
'MsgBox gamepiece(z).status
'w = z + q
'MsgBox w
'z = z + 1


'If gamepiece(z).status <> "Home" Then
'z = z + 1

'MsgBox z

'Exit Do
'Else
'z = z + 1
'End If


'Loop

'Exit For
'End If

'Next



'Open app.path & "\" & gamepiece(ngamepiece).color & " home.txt" For Input As #1

'MsgBox z


'For x = 1 To z
'Line Input #1, c
'Next
'Close #1


'searches = InStr(c, "   ")
'leftss = Mid(c, 1, searches - 1)
'topss = Mid(c, searches, Len(c) - searches + 1)

safeties.leftss = leftss
safeties.topss = topss
gamepiece(ngamepiece).safetyposition = 0
gamepiece(ngamepiece).status = "Home"

Exit Function
End If




If x > 6 Then
'MsgBox x
If Check4.value = 1 Then
MsgBox x
End If


MsgBox "Sorry, you have too much to go home"
rollbacks = True

If remaining = True Then

gamepiece(ngamepiece).boardposition = oldpositions
End If

safeties.leftss = sorrypieces(ngamepiece).Left
safeties.topss = sorrypieces(ngamepiece).Top
Else



'MsgBox x
'MsgBox ngamepiece & "   " & gamepiece(ngamepiece).color


Open App.Path & "\" & gamepiece(ngamepiece).color & " safety.txt" For Input As #1
For z = 1 To x
Line Input #1, c
Next

searches = InStr(c, "  ")
gamepiece(ngamepiece).safetyposition = x
If acceptablemove(ngamepiece, True) = True And continues = True Then


safeties.leftss = Mid(c, 1, searches - 1)
safeties.topss = Mid(c, searches, Len(c) - searhces + 1)
gamepiece(ngamepiece).status = "Safety"
gamepiece(ngamepiece).boardposition = 0
gamepiece(ngamepiece).currentposition = 0
gamepiece(ngamepiece).safetyposition = x
Else
safeties.leftss = previous(1).leftss
safeties.topss = previous(1).topss
gamepiece(ngamepiece).status = previous(1).status
gamepiece(ngamepiece).boardposition = previous(1).boardposition
gamepiece(ngamepiece).currentposition = previous(1).currentposition
gamepiece(ngamepiece).safetyposition = previous(1).safetyposition
End If

Close #1

End If

'x = card(nextcard - 1).card







End Function
Private Function sliding(ngamepiece As Integer, newboards As Boolean, newnumbers, isoptional As Boolean) As positioning
Dim previouspieces() As pieceinfo
ReDim previouspieces(0)
Dim numpieces As Integer
numpieces = 0

'MsgBox newboards

If newboards = True Then
newx = newnumbers - 1
Else
newx = newnumbers
End If
'MsgBox newx & "  for newinfo   " & gamepiece(ngamepiece).boardposition & "   for boardposition"


For x = gamepiece(ngamepiece).boardposition To gamepiece(ngamepiece).boardposition + newx
For z = 1 To UBound(gamepiece)
'MsgBox x
'MsgBox gamepiece(z).boardposition
'MsgBox z
'MsgBox ngamepiece

'MsgBox "Test"

If gamepiece(z).boardposition = x And z <> ngamepiece And gamepiece(z).color = gamepiece(ngamepiece).color And isoptional = True And alreadyasked = False Then
'MsgBox "test again"

ask1 = MsgBox("Are you sure you want to move here because if you do, you will knock your own pawn out", vbYesNo)


If ask1 = 6 Then
alreadyasked = True
End If



If ask1 = 7 Then
continues = False
rollbacks = True
nresetpieces numpieces, previouspieces()
sliding.leftss = previous(1).leftss
sliding.topss = previous(1).topss

For xx = 1 To UBound(previous)

newpiece = previous(xx).gamepiece

gamepiece(newpiece).boardposition = previous(xx).boardposition
'MsgBox gamepiece(ngamepiece).boardposition

gamepiece(newpiece).currentposition = previous(xx).currentposition
gamepiece(newpiece).status = previous(xx).status
gamepiece(newpiece).safetyposition = previous(xx).safetyposition

Next


If opponent.gamepiece <> 0 Then

newx = opponent.gamepiece
gamepiece(newx).boardposition = opponent.boardposition
gamepiece(newx).safetyposition = opponent.safetyposition
gamepiece(newx).currentposition = opponent.currentposition
gamepiece(newx).status = opponent.status
gamepiece(newx).safetyposition = opponent.safetyposition
sorrypieces(newx).Left = opponent.leftss
sorrypieces(newx).Top = opponent.topss
'MsgBox "Test this"
'MsgBox newx
'MsgBox ngamepiece

End If




Exit Function
Exit For
Exit For
End If




End If


If gamepiece(z).boardposition = x And z <> ngamepiece Then
numpieces = numpieces + 1
ReDim Preserve previouspieces(numpieces)
previouspieces(numpieces).gamepiece = z
previouspieces(numpieces).color = gamepiece(z).color
previouspieces(numpieces).currentposition = gamepiece(z).currentposition
previouspieces(numpieces).boardposition = gamepiece(z).boardposition
previouspieces(numpieces).leftss = sorrypieces(z).Left
previouspieces(numpieces).topss = sorrypieces(z).Top
previouspieces(numpieces).safetyposition = gamepiece(z).safetyposition
previouspieces(numpieces).status = gamepiece(z).status
'MsgBox "Test again"
'MsgBox z & "  test"


backtostart (z)
Exit For
End If
Next
Next



If newboards = True Then
gamepiece(ngamepiece).boardposition = 1


For z = 1 To UBound(gamepiece)


If gamepiece(z).boardposition = x And z <> ngamepiece And gamepiece(z).color = gamepiece(ngamepiece).color And isoptional = True And alreadyasked = False Then
ask1 = MsgBox("Are you sure you want to move here because if you do, you will knock your own pawn out", vbYesNo)
If ask1 = 6 Then
alreadyasked = True
End If
If ask1 = 7 Then
nresetpieces numpieces, previouspieces()
sliding.leftss = previous(1).leftss
sliding.topss = previous(1).topss
continues = False
rollbacks = True
For xx = 1 To UBound(previous)

newpiece = previous(xx).gamepiece


gamepiece(newpiece).boardposition = previous(xx).boardposition
gamepiece(newpiece).currentposition = previous(xx).currentposition
gamepiece(newpiece).status = previous(xx).status
gamepiece(newpiece).safetyposition = previous(xx).safetyposition
Next

If opponent.gamepiece <> 0 Then
newx = opponent.gamepiece
gamepiece(newx).boardposition = opponent.boardposition
gamepiece(newx).safetyposition = opponent.safetyposition
gamepiece(newx).currentposition = opponent.currentposition
gamepiece(newx).status = opponent.status
gamepiece(newx).safetyposition = opponent.safetyposition
sorrypieces(newx).Left = opponent.leftss
sorrypieces(newx).Top = opponent.topss
End If
Exit Function
Exit For
Exit For
End If




End If


If gamepiece(z).boardposition = 1 And z <> ngamepiece Then
numpieces = numpieces + 1
ReDim Preserve previouspieces(numpieces)
previouspieces(numpieces).gamepiece = z



previouspieces(numpieces).color = gamepiece(z).color
previouspieces(numpieces).currentposition = gamepiece(z).currentposition
previouspieces(numpieces).boardposition = gamepiece(z).boardposition
previouspieces(numpieces).leftss = sorrypieces(z).Left
previouspieces(numpieces).topss = sorrypieces(z).Top
previouspieces(numpieces).safetyposition = gamepiece(z).safetyposition
previouspieces(numpieces).status = gamepiece(z).status
backtostart (z)
'MsgBox z

Exit For
End If
Next

Else
gamepiece(ngamepiece).boardposition = gamepiece(ngamepiece).boardposition + newnumbers
End If

gamepiece(ngamepiece).currentposition = gamepiece(ngamepiece).currentposition + newnumbers
newx = gamepiece(ngamepiece).boardposition - 1
If newx < 0 Then
newx = 0
End If


searches = InStr(List1.List(newx), " ")
sliding.leftss = Mid(List1.List(newx), 1, searches - 1)
sliding.topss = Mid(List1.List(newx), searches, Len(List1.List(newx)) - searches + 1)
End Function
Private Sub nresetpieces(numpieces As Integer, previouspieces() As pieceinfo)
If numpieces <> 0 Then
'MsgBox numpieces

For x = 1 To numpieces
newx = previouspieces(x).gamepiece
gamepiece(newx).currentposition = previouspieces(x).currentposition
sorrypieces(newx).Left = previouspieces(x).leftss
sorrypieces(newx).Top = previouspieces(x).topss
gamepiece(newx).status = previouspieces(x).status
gamepiece(newx).safetyposition = previouspieces(x).safetyposition
gamepiece(newx).boardposition = previouspieces(x).boardposition

Next
End If


'previouspieces(numpieces).gamepiece = z
'previouspieces(numpieces).color = gamepiece(z).color
'previouspieces(numpieces).currentposition = gamepiece(z).currentposition
'previouspieces(numpieces).boardposition = gamepiece(z).boardposition
'previouspieces(numpieces).leftss = sorrypieces(z).Left
'previouspieces(numpieces).topss = sorrypieces(z).Top
'previouspieces(numpieces).safetyposition = gamepiece(z).safetyposition
'previouspieces(numpieces).status = gamepiece(z).status

End Sub
Private Function newposition(ngamepiece As Integer, getpiece, values) As positioning
'searches = InStr(List1.List(0), " ")
'leftss = Mid(List1.List(0), 1, searches - 1)
'topss = Mid(List1.List(0), searches, Len(c) - searches + 1)
Dim newx As Integer
'MsgBox "Test"
Dim oldpositions As Integer

oldpositions = gamepiece(ngamepiece).boardposition

If card(nextcard - 1).card = 13 Then
'temp only


newx = gamepiece(getpiece).boardposition
gamepiece(ngamepiece).boardposition = gamepiece(getpiece).boardposition
'Open "C:\E.txt" For Output As #1
gamepiece(ngamepiece).currentposition = currentposition(ngamepiece, getpiece)
gamepiece(ngamepiece).status = "Out"

'For x = 0 To 59
'z = x + gamepiece(ngamepiece).starting
'If z > 59 Then
'z = z - 60
'End If

'searches = InStr(List1.List(z), " ")
'leftss = Mid(List1.List(z), 1, searches - 1)
'topss = Mid(List1.List(z), searches, Len(List1.List(z)) - searhces + 1)
'MsgBox sorrypieces(getpiece).Left & "   " & leftss & vbCrLf & sorrypieces(getpiece).Top & "   " & topss & vbCrLf & z

'MsgBox z

'Write #1, z



'If sorrypieces(getpiece).Left = leftss And sorrypieces(getpiece).Top = topss Then

'gamepiece(ngamepiece).currentposition = x + 1
'gamepiece(ngamepiece).status = "Out"
'MsgBox gamepiece(ngamepiece).currentposition
'MsgBox z
'MsgBox gamepiece(ngamepiece).currentposition & "   " & x & "   " & gamepiece(ngamepiece).starting



'MsgBox z
'MsgBox leftss
'MsgBox topss
newposition.leftss = sorrypieces(getpiece).Left
newposition.topss = sorrypieces(getpiece).Top

backtostart (getpiece)
'Exit For
'End If
'Next
'Close #1

'newposition.leftss = leftss
'newposition.topss = topss
'MsgBox leftss
'MsgBox topss

Exit Function
End If

'MsgBox "Test so far"


If gamepiece(ngamepiece).status = "Safety" Then
newposition = safeties(ngamepiece, 0, False, 0)

Exit Function
End If



If gamepiece(ngamepiece).status = "Start" Then
newx = gamepiece(ngamepiece).starting
gamepiece(ngamepiece).boardposition = newx

ElseIf values <> 0 Then
newx = gamepiece(ngamepiece).boardposition + values
'MsgBox newx

gamepiece(ngamepiece).boardposition = newx
gamepiece(ngamepiece).status = "Out"
ElseIf card(nextcard - 1).card = 10 And gamepiece(ngamepiece).currentposition > 54 And gamepiece(ngamepiece).currentposition <> 59 Then
newx = gamepiece(ngamepiece).boardposition - 1
gamepiece(ngamepiece).boardposition = newx
gamepiece(ngamepiece).status = "Out"
ElseIf card(nextcard - 1).card = 4 Then
newx = gamepiece(ngamepiece).boardposition - 4
gamepiece(ngamepiece).boardposition = newx
gamepiece(ngamepiece).status = "Out"
ElseIf card(nextcard - 1).card = 13 Then
newx = gamepiece(ngamepiece).boardposition

Else

newx = gamepiece(ngamepiece).boardposition + card(nextcard - 1).card
gamepiece(ngamepiece).boardposition = newx
gamepiece(ngamepiece).status = "Out"

End If
'MsgBox newx
'MsgBox valuess

If newx > List1.ListCount Then
newx = newx - List1.ListCount
gamepiece(ngamepiece).boardposition = newx
End If
If newx < 1 Then
newx = 60 + newx



End If
'MsgBox gamepiece(ngamepiece).currentposition

If values = 0 Then



If card(nextcard - 1).card = 10 And gamepiece(ngamepiece).currentposition > 54 And gamepiece(ngamepiece).currentposition <> 59 Then
newcard = -1

ElseIf gamepiece(ngamepiece).status = "Start" Then
'gamepiece(ngamepiece).status = "Out"
newcard = 0

ElseIf card(nextcard - 1).card = 13 Then
newcard = 0
ElseIf card(nextcard - 1).card = 4 Then
newcard = -4
Else
newcard = card(nextcard - 1).card
End If
'MsgBox newcard
Else
newcard = values
End If
'MsgBox newcard
If Check4.value = 1 Then
MsgBox gamepiece(ngamepiece).currentposition & " for current position " & newcard & " for newcard " & gamepiece(ngamepiece).currentposition + newcard & " for new current position"

End If


If gamepiece(ngamepiece).currentposition + newcard > 58 And gamepiece(ngamepiece).currentposition <> 59 Then

'MsgBox gamepiece(ngamepiece).currentposition + newx
If gamepiece(ngamepiece).status = "Start" Then
gamepiece(ngamepiece).status = "Out"
End If
Dim newinfo As Integer

newinfo = gamepiece(ngamepiece).currentposition + newcard - 59


newposition = safeties(ngamepiece, newinfo, True, oldpositions)


'MsgBox "You will be safe"
'gamepiece(ngamepiece).boardposition = oldpositions

'newposition.leftss = sorrypieces(ngamepiece).Left
'newposition.topss = sorrypieces(ngamepiece).Top
Exit Function
End If
'MsgBox currentposition

If gamepiece(ngamepiece).currentposition = 59 And newcard < 1 Then

gamepiece(ngamepiece).currentposition = gamepiece(ngamepiece).currentposition + newcard
ElseIf gamepiece(ngamepiece).currentposition = 59 Then
gamepiece(ngamepiece).currentposition = newcard - 1
Else



gamepiece(ngamepiece).currentposition = gamepiece(ngamepiece).currentposition + newcard
If gamepiece(ngamepiece).currentposition < 1 And gamepiece(ngamepiece).status <> "Start" Then
gamepiece(ngamepiece).currentposition = 60 + gamepiece(ngamepiece).currentposition
End If
If gamepiece(ngamepiece).status = "Start" Then
gamepiece(ngamepiece).status = "Out"
End If
End If

'If gamepiece(ngamepiece).currentposition > 58 Then
'MsgBox card(nextcard - 1).card

'gamepiece(ngamepiece).currentposition = card(nextcard - 1).card
'End If
'MsgBox newx
'MsgBox gamepiece(ngamepiece).currentposition
'fix for split move try also for other moves if this still has a problem

'newx = gamepiece(ngamepiece).boardposition

If gamepiece(ngamepiece).boardposition < 1 Then
gamepiece(ngamepiece).boardposition = newx
End If



newx = newx - 1
'MsgBox newx


searches = InStr(List1.List(newx), " ")
If acceptablemove(ngamepiece, False) = True Then

newposition.leftss = Mid(List1.List(newx), 1, searches - 1)
newposition.topss = Mid(List1.List(newx), searches, Len(List1.List(newx)) - searches + 1)
Else
newposition.leftss = previous(1).leftss
newposition.topss = previous(1).topss
End If



'MsgBox gamepiece(ngamepiece).currentposition




End Function
Private Function acceptablemove(ngamepiece As Integer, safeties As Boolean) As Boolean
acceptablemove = True
If Check2.value = 1 Then
MsgBox safeties
End If

'MsgBox safeties



For z = 1 To UBound(gamepiece)
If safeties = False Then
firstposition = gamepiece(z).boardposition
secondposition = gamepiece(ngamepiece).boardposition
Else
firstposition = gamepiece(z).safetyposition
secondposition = gamepiece(ngamepiece).safetyposition
End If




If gamepiece(z).status = gamepiece(ngamepiece).status And firstposition = secondposition And z <> ngamepiece And gamepiece(z).color = gamepiece(ngamepiece).color And alreadyasked = False Then

'MsgBox z & "   number on " & ngamepiece & "   number " & gamepiece(ngamepiece).status & "  " & gamepiece(z).status

ask1 = MsgBox("Are you sure you want to do this move.  If you do, you will knock yourself out", vbYesNo)
If ask1 = 6 Then
alreadyasked = True
End If

If ask1 = 7 Then
acceptablemove = False
continues = False
rollbacks = True
'acceptablemove.leftss = previous(1).leftss
'newposition.topss = previous(1).topss
'MsgBox uboundprevious

For xx = 1 To UBound(previous)
newpiece = previous(xx).gamepiece
'If previous(xx).status <> "Start" Then


gamepiece(newpiece).boardposition = previous(xx).boardposition
gamepiece(newpiece).currentposition = previous(xx).currentposition
gamepiece(newpiece).status = previous(xx).status
gamepiece(newpiece).safetyposition = previous(xx).safetyposition
'End If

Next


Exit Function
Exit For
End If
End If





'If firstposition = secondposition And z <> ngamepiece And gamepiece(z).status = gamepiece(ngamepiece).status
If firstposition = secondposition And z <> ngamepiece And safeties = True And gamepiece(z).color = gamepiece(ngamepiece).color Then



backtostart (z)
Exit For
End If
If firstposition = secondposition And z <> ngamepiece And safeties = False And gamepiece(z).status = gamepiece(ngamepiece).status Then
backtostart (z)
Exit For
End If



Next
If Check4.value = 1 Then
MsgBox "Continue after acceptable move function"
End If

End Function
Private Function moveouts() As moveoutinfo
Dim pawntaken As Integer
moveouts.willmove = False
If Check4.value = 1 Then
MsgBox "Continue"
End If


Dim optionalss As Boolean
optionalss = False
'z = 0
pawntaken = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Start" Then
pawntaken = x
Exit For
End If
Next
If pawntaken <> 0 Then

For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).currentposition = 0 And gamepiece(x).status = "Out" Then
optionalss = True
Exit For
End If
Next
End If



If pawntaken <> 0 And optionalss = False Or pawntaken <> 0 And card(nextcard - 1).card = 13 Then


moveouts.willmove = True
moveouts.pawn = pawntaken
If Check4.value = 1 Then

MsgBox pawntaken
End If


End If







End Function
Private Function possiblemoves() As moveinfo
If unloadthis = True Then
On Error Resume Next
End If

Dim newplay() As possiblemove
Dim z As Integer
Dim newnumbers As Integer


ReDim newplay(0)
Dim requiress As String
If card(nextcard - 1).card = 4 Then
newnumbers = -4
Else
newnumbers = card(nextcard - 1).card
End If

Dim currentsafe As Boolean
currentsafe = False
z = 0


'MsgBox unloadthis
For x = players(turnnum).startnum To players(turnnum).startnum + 3

If gamepiece(x).status = "Out" Or gamepiece(x).status = "Safety" Then


If gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 5 Then
newleft = 1
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 4 Then
newleft = 2
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 3 Then
newleft = 3
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 2 Then
newleft = 4
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 1 Then
newleft = 5
ElseIf gamepiece(x).status = "Out" And gamepiece(x).currentposition = 59 Then
newleft = 59 + 6
'MsgBox x

Else
newleft = 58 - gamepiece(x).currentposition + 6
End If
If gamepiece(x).status = "Safety" Then
currentsafe = True
Else
currentsafe = False

End If
'MsgBox gamepiece(x).currentposition


If CInt(newleft) >= CInt(newnumbers) Then
z = z + 1
ReDim Preserve newplay(z)
newplay(z).safety = currentsafe
currentsafe = False
newplay(z).playnum = newnumbers

newplay(z).pawn = x
End If

End If
Next
If z = 0 Then
possiblemoves.optional = 0
possiblemoves.required = 0
Exit Function
End If








'searches = InStr(List1.List(0), " ")
'leftss = Mid(List1.List(0), 1, searches - 1)
'topss = Mid(List1.List(0), searches, Len(c) - searches + 1)
'Dim newx As Integer
'newx = newnumbers
'Dim continueprocess As Boolean
Dim latestpawn As Integer
If Check4.value = 1 Then
MsgBox UBound(newplay)
End If

For z = 1 To UBound(newplay)
'continueprocess = True
'requiress = True

'latestpawn = newplay(z).pawn
requiress = "Yes"
requiress = processpossibleplays(z, newplay())
If requiress = "Yes" Then


'latestpawn = newplay(1, z)
'MsgBox newplay(z).pawn

latestpawn = newplay(z).pawn
latestvalue = newplay(z).playnum
'MsgBox latestpawn


'latestpawn = newplay(1, z)
'newnumbers = newplay(2, z)
'newnumbers = newplay(z).playnum
possiblemoves.required = possiblemoves.required + 1
ElseIf requiress = "Optional" Then
possiblemoves.optional = possiblemoves.optional + 1
End If


Next


If UBound(newplay) = 1 Then
possiblemoves.firstpawn = newplay(1).pawn

possiblemoves.value = newplay(1).playnum
ElseIf possiblemoves.required = 1 Then
possiblemoves.firstpawn = latestpawn
possiblemoves.value = latestvalue

End If
'MsgBox possiblemoves.firstpawn

Exit Function




















End Function
Private Function requireslide()

End Function
Private Function requirethismove(safety, newpositions) As Boolean
requirethismove = True
If safety = 6 Then
requirethismove = True
Exit Function
End If

If safety <> 0 Then
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).safetyposition = safety Then
requirethismove = False
Exit Function
End If
Next

Exit Function
End If






'possiblemoves.required = possiblemoves.required + 1
'continueprocess = False
'End If


For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).currentposition = newpositions And gamepiece(x).status = "Out" Then

requirethismove = False
'MsgBox x
'MsgBox gamepiece(x).currentposition
'MsgBox newpositions

Exit Function
Exit For
End If

'If gamepiece(x).currentposition <> newpositions Then
'sorrymoves.optional = sorrymoves.optional + 1
'Else
'sorrymoves.required = sorrymoves.required + 1
'End If
Next




End Function
Private Function processpossibleplays(z As Integer, newplay() As possiblemove) As String
processpossibleplays = "Yes"
Dim requiress As Boolean
Dim slidess As Boolean
Dim possiblehits As Boolean
slidess = False
possiblehits = False
Dim latestpawn As Integer

requiress = True

'latestpawn = newplay(1, z)
'MsgBox x
latestpawn = newplay(z).pawn

'latestpawn = newplay(1, z)
'newnumbers = newplay(2, z)
newnumbers = newplay(z).playnum

Dim oldposition As Integer
oldposition = gamepiece(newplay(z).pawn).currentposition
'Dim newposition As Integer



'If gamepiece(ngamepiece).currentposition = 59 Then
If oldposition = 59 Then


newpositions = newnumbers - 1
Else



newpositions = oldposition + newnumbers

If oldposition < 1 And oldposition <> 0 Then
newpositions = 60 + oldposition
ElseIf gamepiece(newplay(z).pawn).status = "Safety" And newnumbers = -1 And gamepiece(newplay(z).pawn).safetyposition = 1 Then
newpositions = 58
'currentsafe = False
newplay(z).safety = False

End If

End If


If oldposition < 1 And oldposition <> 0 Then
newpositions = 60 + oldposition

ElseIf gamepiece(newplay(z).pawn).status = "Safety" And newnumbers = -4 And gamepiece(newplay(z).pawn).safetyposition = 4 Then
newpositions = 58
currentsafe = False

ElseIf gamepiece(newplay(z).pawn).status = "Safety" And newnumbers = -4 And gamepiece(newplay(z).pawn).safetyposition = 3 Then
newpositions = 57
currentsafe = False
ElseIf gamepiece(newplay(z).pawn).status = "Safety" And newnumbers = -4 And gamepiece(newplay(z).pawn).safetyposition = 2 Then
newpositions = 56
currentsafe = False
ElseIf gamepiece(newplay(z).pawn).status = "Safety" And newnumbers = -4 And gamepiece(newplay(z).pawn).safetyposition = 1 Then
newpositions = 55
currentsafe = False


'newpositions = 58
End If



'End If
If newnumbers = -4 And oldposition < 4 And gamepiece(newplay(z).pawn).status = "Out" Then
'newpositions = ""
newpositions = oldposition + 56
End If
'MsgBox oldposition



If newpositions > 59 And oldposition <> 59 Or newplay(z).safety = True Then




'If gamepiece(ngamepiece).currentposition + newcard > 58 And gamepiece(ngamepiece).currentposition <> 59 Then

'MsgBox gamepiece(ngamepiece).currentposition + newx
If newplay(z).safety = False Then
newinfo = oldposition + newnumbers - 58
Else
newinfo = gamepiece(newplay(z).pawn).safetyposition + newnumbers
End If
If Check4.value = 1 Then
MsgBox newinfo & " for newinfo " & oldposition & " for old position " & newnumbers & " for number used for calculation"
End If


'MsgBox z

'newinfo = oldposition + newnumbers - 59


'newposition = safeties(ngamepiece, newinfo, True, oldpositions)


requiress = requirethismove(newinfo, newinfo)


If requiress = True Then
'processpossibleplays = "No"
processpossibleplays = "Yes"
Else
processpossibleplays = "No"
End If


Exit Function

'possible10.required = possible10.required + 1
'pawns = latestpawn

End If

'continueprocess = False
'End If





'If continueprocess = True Then






'Dim slidess As Boolean
'Dim possiblehits As Boolean
possiblehits = False
slidess = False


slidess = sliderule(latestpawn, False, False)


If slidess = True Then
continueprocess = False
possiblehits = possiblehitslide(latestpawn)

Else

possiblehits = False
End If


If possiblehits = False And slidess = True Then
'possible10.required = possible10.required + 1
'pawns = latestpawn
processpossibleplays = "Yes"
Exit Function
End If

If possiblehits = True And slidess = True Then

processpossibleplays = "Optional"
Exit Function

'possible10.optional = possible10.optional + 1


End If
If Check4.value = 1 Then
MsgBox newpositions
End If

requiress = requirethismove(0, newpositions)

'For x = players(turnnum).startnum To players(turnnum).startnum + 3
'If gamepiece(x).currentposition = newpositions Then
'requiress = False
'Exit For
'End If

'If gamepiece(x).currentposition <> newpositions Then
'sorrymoves.optional = sorrymoves.optional + 1
'Else
'sorrymoves.required = sorrymoves.required + 1
'End If
'Next
If requiress = True Then
'possible10.required = possible10.required + 1
'pawns = latestpawn
processpossibleplays = "Yes"
Else
'MsgBox newpositions
processpossibleplays = "No"

End If
'End If



End Function
Private Function possible10() As moveinfo
Dim z As Integer
Dim newnumbers As Integer


Dim currentsafe As Boolean
currentsafe = False
Dim newplay() As possiblemove
ReDim newplay(0)

'ReDim newplay(2, 0)
Dim requiress As String

'If card(nextcard - 1).card = 4 Then
'newnumbers = -4
'Else
newnumbers = card(nextcard - 1).card
'End If
'gamepiece(newplay(1, z))
z = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3

If gamepiece(x).status = "Out" Or gamepiece(x).status = "Safety" Then
If gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 5 Then
newleft = 1
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 4 Then
newleft = 2
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 3 Then
newleft = 3
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 2 Then
newleft = 4
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 1 Then
newleft = 5
ElseIf gamepiece(x).status = "Out" And gamepiece(x).currentposition = 59 Then
newleft = 59 + 6
Else
newleft = 58 - gamepiece(x).currentposition + 6
End If
If Check4.value = 1 Then
MsgBox newleft & "   " & gamepiece(x).currentposition
End If



If newleft < newnumbers Then
playnumber = -1
Else
playnumber = newnumbers
End If
If gamepiece(x).status = "Safety" Then
currentsafe = True
Else
currentsafe = False
End If


If CInt(newleft) >= CInt(playnumber) Then
z = z + 1
ReDim Preserve newplay(z)
newplay(z).safety = currentsafe
currentsafe = False
newplay(z).pawn = x
newplay(z).playnum = playnumber

'newplay(1, z) = x
'newplay(2, z) = playnumber
'newplay(3, z) = currentsafe

End If

End If
Next
If z = 0 Then
possible10.optional = 0
possible10.required = 0
Exit Function
End If








'searches = InStr(List1.List(0), " ")
'leftss = Mid(List1.List(0), 1, searches - 1)
'topss = Mid(List1.List(0), searches, Len(c) - searches + 1)
'Dim newx As Integer
'newx = newnumbers

'Dim continueprocess As Boolean
Dim latestpawn As Integer
Dim newinfo As Integer

For z = 1 To UBound(newplay)

'continueprocess = True
'status = "Required"
requiress = "Yes"

'check
'newnumbers = newplay(z).playnum

requiress = processpossibleplays(z, newplay())
'If requiress = True Then
If requiress = "Yes" Then


'latestpawn = newplay(1, z)
'MsgBox x
latestpawn = newplay(z).pawn
latestvalue = newplay(z).playnum


'latestpawn = newplay(1, z)
'newnumbers = newplay(2, z)
'newnumbers = newplay(z).playnum
possible10.required = possible10.required + 1
ElseIf requiress = "Optional" Then
possible10.optional = possible10.optional + 1
End If




'Next




If newplay(z).playnum <> "-1" Then
'newnumbers = -1
newplay(z).playnum = -1

requiress = "Yes"
requiress = processpossibleplays(z, newplay())

If requiress = "Yes" Then


'latestpawn = newplay(1, z)
'MsgBox x
latestpawn = newplay(z).pawn
latestvalue = newplay(z).playnum


'latestpawn = newplay(1, z)
'newnumbers = newplay(2, z)
'newnumbers = newplay(z).playnum
possible10.required = possible10.required + 1
ElseIf requiress = "Optional" Then
possible10.optional = possible10.optional + 1
End If
End If

Next





If UBound(newplay) = 1 Then
possible10.firstpawn = newplay(1).pawn
possible10.value = newplay(1).playnum

ElseIf possible10.required = 1 Then
possible10.firstpawn = latestpawn
possible10.value = latestvalue



End If
Exit Function



End Function
Private Function sorrymoves() As moveinfo
'slidenum
sorrymoves.optional = 0
sorrymoves.required = 0

Dim latestpawn As Integer
Dim slidess As Boolean
Dim possiblehits As Boolean
Dim pawntaken As Integer
z = 0
pawntaken = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Start" Then


z = z + 1
If pawntaken = 0 Then
pawntaken = x
End If

End If
Next


If z = 0 Then
possiblemoves.required = 0
possiblemoves.optional = 0
Exit Function
End If



z = 0
For x = 1 To UBound(gamepiece)
If gamepiece(x).color <> color.Caption And gamepiece(x).status = "Out" Then
z = z + 1
latestpawn = x
slidess = sliderule(latestpawn, False, False)

If slidess = True Then
possiblehits = possiblehitslide(latestpawn)
Else
possiblehits = False
End If

If possiblehits = False Then
sorrymoves.required = sorrymoves.required + 1
Else
sorrymoves.optional = sorrymoves.optional + 1
End If




End If
Next
If z = 0 Then
sorrymoves.required = 0
sorrymoves.optional = 0
Exit Function
End If



If z = 1 Then
sorrymoves.firstpawn = pawntaken
sorrymoves.secondpawn = latestpawn
End If


End Function
Private Function possiblehitslide(newpawn As Integer) As Boolean
'If newboards = True Then
'If slidenum = 58 Then
If gamepiece(newpawn).boardposition = 58 Then
newx = slidenum - 1
Else
newx = slidenum
End If
'58
For x = gamepiece(newpawn).boardposition To gamepiece(newpawn).boardposition + newx


For z = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(z).boardposition = x Then
possiblehitslide = True
Exit Function
Exit For
Exit For
End If
Next
Next
x = 1

For z = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(z).boardposition = x Then
possiblehitslide = True
Exit Function
Exit For
End If
Next
possiblehitslide = False





End Function
Private Function possible11() As Integer
'MsgBox "Check so far"

z = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Out" Then
z = z + 1
End If
Next
If z = 0 Then
possible11 = 0
Exit Function
End If

firstpossible = z



z = 0
For x = 1 To UBound(gamepiece)
If gamepiece(x).color <> players(turnnum).color And gamepiece(x).status = "Out" Then

z = z + 1
End If
Next
If z = 0 Then
possible11 = 0
Else

possible11 = firstpossible * z
End If




End Function

Private Function possible7() As moveinfo
'MsgBox "Test so far"


Dim newplay() As possiblemove
Dim z As Integer
Dim newnumbers As Integer


ReDim newplay(0)
Dim requiress As String
'If card(nextcard - 1).card = 4 Then
'newnumbers = -4
'Else
'newnumbers = card(nextcard - 1).card
'End If

newnumbers = 7

Dim currentsafe As Boolean
currentsafe = False
z = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Out" Or gamepiece(x).status = "Safety" Then


If gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 5 Then
newleft = 1
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 4 Then
newleft = 2
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 3 Then
newleft = 3
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 2 Then
newleft = 4
ElseIf gamepiece(x).status = "Safety" And gamepiece(x).safetyposition = 1 Then
newleft = 5
ElseIf gamepiece(x).status = "Out" And gamepiece(x).currentposition = 59 Then
newleft = 59 + 6
'MsgBox x

Else
newleft = 58 - gamepiece(x).currentposition + 6
End If
If gamepiece(x).status = "Safety" Then
currentsafe = True
End If
'MsgBox gamepiece(x).currentposition


'If CInt(newleft) > CInt(newnumbers) Then
z = z + 1
ReDim Preserve newplay(z)
newplay(z).safety = currentsafe
currentsafe = False
newplay(z).playnum = newnumbers
newplay(z).leftss = newleft

newplay(z).pawn = x
'End If

End If
Next
If z = 0 Then
possible7.optional = 0
possible7.required = 0
Exit Function
ElseIf z = 1 And CInt(newplay(1).leftss) <= CInt(newplay(1).playnum) Then
possible7.optional = 0
possible7.required = 0
'MsgBox newplay(1).leftss
'MsgBox newplay(1).playnum



Exit Function
End If




If z = 1 Then
possible7 = possiblemoves
Exit Function
End If
If z = 2 And CInt(newplay(1).leftss) + CInt(newplay(2).leftss) = 7 Then
possible7.optional = 0
possible7.required = 1
possible7.firstpawn = newplay(1).pawn
possible7.secondpawn = newplay(2).pawn
possible7.value = newplay(1).leftss
Exit Function
End If
If z = 2 And CInt(newplay(1).leftss) + CInt(newplay(2).leftss) < 7 Then
possible7.optional = 0
possible7.required = 0
Exit Function
End If
If z = 2 And CInt(newplay(1).leftss) < 7 And CInt(newplay(2).leftss) < 9 Then

possible7.optional = 0
possible7.required = 1
possible7.firstpawn = newplay(1).pawn
possible7.secondpawn = newplay(2).pawn
possible7.value = newplay(1).leftss
Exit Function
End If
If z = 2 And CInt(newplay(2).leftss) < 7 And CInt(newplay(1).leftss) < 9 Then

possible7.optional = 0
possible7.required = 1
possible7.firstpawn = newplay(2).pawn
possible7.secondpawn = newplay(1).pawn
possible7.value = newplay(2).leftss
Exit Function
End If





possible7.optional = 0
possible7.required = 2




End Function
Private Sub processthismove()
'MsgBox players(playnum).startnum
'MsgBox playnum
'MsgBox "Test so far"
rollbacks = False
Dim winnings As Boolean

If unloadthis = True Then
On Error Resume Next
End If

disableall
todolabel.Visible = False
todo.Visible = False
todo.Caption = ""


status.Caption = "Figuring out possible moves"

'Dim willmoveout As moveoutinfo
willmoveout.willmove = False
willmoveout.pawn = 0

Dim optionals As Integer



slidenum = 0
Dim newmove As moveinfo
If card(nextcard - 1).card = 13 Then
willmoveout = moveouts
newmove = sorrymoves



ElseIf card(nextcard - 1).card = 1 Or card(nextcard - 1).card = 2 Then
willmoveout = moveouts



End If

If card(nextcard - 1).card = 1 Or card(nextcard - 1).card = 2 Or card(nextcard - 1).card = 3 Or card(nextcard - 1).card = 4 Or card(nextcard - 1).card = 5 Or card(nextcard - 1).card = 8 Or card(nextcard - 1).card = 12 Then
'nextmoves = newmove.firstpawn
newmove = possiblemoves



If willmoveout.willmove = True Then

newmove.required = newmove.required + 1
End If
If newmove.required = 1 And willmoveout.willmove = True Then
'newmove.firstpawn = nextmoves
newmove.firstpawn = willmoveout.pawn



ElseIf newmove.required <> 1 Then
newmove.firstpawn = 0
newmove.secondpawn = 0
End If





ElseIf card(nextcard - 1).card = 10 Then
newmove = possible10
If newmove.required <> 1 Then
newmove.firstpawn = 0
newmove.secondpawn = 0
End If



ElseIf card(nextcard - 1).card = 11 Then
optionals = possible11
newmove = possiblemoves

newmove.optional = newmove.optional + optionals
If newmove.optional <> 0 Then
newmove.firstpawn = 0
newmove.secondpawn = 0
End If
If optionals = 0 Then
Check1.Enabled = False
Check1.value = 0
Else
Check1.Enabled = True
End If


ElseIf card(nextcard - 1).card = 7 Then
newmove = possible7
If newmove.required <> 1 Then
newmove.firstpawn = 0
newmove.secondpawn = 0
End If
'If newmove.required > 1 Then





End If

If Check4.value = 1 Then
MsgBox newmove.firstpawn & "  first pawn " & vbCrLf & newmove.secondpawn & " second pawn " & vbCrLf & newmove.optional & " for optional " & vbCrLf & newmove.required & " for required for sorry card" & vbCrLf & newmove.value & " for amount to move"
End If
If newmove.optional <> 0 And newmove.required = 0 Then

pass.Enabled = True

Else
pass.Enabled = False
End If
'this is used for testing  please leave
Dim endturns As Boolean
If newmove.required <> 1 Or newmove.optional <> 0 Then
status.Caption = "Waiting for user input"

If newmove.required > 1 Or newmove.optional <> 0 Then




firstprocessmanuel
firstprocessmanuel

'MsgBox newmove.firstpawn & "  first pawn " & vbCrLf & newmove.secondpawn & " second pawn " & vbCrLf & newmove.optional & " for optional " & vbCrLf & newmove.required & " for required for sorry card" & vbCrLf & newmove.value & " for amount to move"
Exit Sub


End If

End If

If newmove.required = 1 And newmove.optional = 0 Then

'MsgBox newmove.firstpawn & "  first pawn " & vbCrLf & newmove.secondpawn & " second pawn " & vbCrLf & newmove.optional & " for optional " & vbCrLf & newmove.required & " for required for sorry card" & vbCrLf & newmove.value & " for amount to move"
status.Caption = "Taking turn"
If Check4.value = 1 Then
MsgBox newmove.value
End If

endturns = taketurn(newmove.firstpawn, newmove.secondpawn, newmove.value, True)
If testing = False Then

status.Caption = "Taking break so user can see what happened"

Pause (secondss)

End If

If unloadthis = True Then
Exit Sub
End If


'MsgBox endturns
End If

'end test
'For x = players(turnnum).startnum To players(turnnum).startnum + 3
'MsgBox x
'Next



If endturns = True And card(nextcard - 1).card <> 2 Then
winnings = haswon
If winnings = True Then
Unload Me
Exit Sub
End If

newturn
drawcards

If testing = False Then
processthismove
End If

ElseIf endturns = True And card(nextcard - 1).card = 2 Then
status.Caption = "Taking additional turn"
winnings = haswon
If winnings = True Then
Unload Me
Exit Sub
End If

'newturn
drawcards

If testing = False Then
processthismove
End If




ElseIf newmove.optional = 0 And newmove.required = 0 Then
If testing = False Then

status.Caption = "Taking break so user can see what happened"
Pause (secondss)
End If

If unloadthis = True Then
Exit Sub
End If

winnings = haswon
If winnings = True Then
Unload Me
Exit Sub
End If

newturn
drawcards

If testing = False Then
processthismove
End If




ElseIf card(nextcard - 1).card <> 2 And newmove.optional = 0 Then
'MsgBox newmove.required & "  required  " & newmove.optional & "   optional"

'firstprocessmanuel

'newturn
status.Caption = "Waiting for user input"

End If
End Sub

Private Sub Command7_Click()
newturn
drawcards
End Sub

Private Sub Command8_Click()
rollback.Enabled = True

processthismove






End Sub
Sub Pause(interval)  'Pause an interval

    current = Timer
    'MsgBox Timer
    
    Do While Timer - current < Val(interval)
       If unloadthis = True Then
       Exit Sub
       Exit Do
       End If
       
    'MsgBox Timer - current
    
        
        DoEvents
    Loop
    
End Sub
Private Sub newturn()




On Error Resume Next

status.Caption = "Ending turn"






turnnum = turnnum + 1
If turnnum > UBound(players) Then
turnnum = 1
End If
player.Caption = players(turnnum).playername
color.Caption = players(turnnum).color
End Sub
Private Function haswon() As Boolean
haswon = False

z = 0
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Home" Then
z = z + 1
End If
Next
If z = 4 Then
MsgBox player.Caption & " with color " & color.Caption & " has won because he/she got all 4 men home"

haswon = True

'unloadthis = True
'Unload Me
Exit Function
End If



End Function
Private Sub previousturns()




On Error Resume Next

status.Caption = "Rolling Back turn"






turnnum = turnnum - 1
If turnnum = 0 Then
turnnum = UBound(players)
End If
player.Caption = players(turnnum).playername
color.Caption = players(turnnum).color
End Sub



Private Function taketurn(firstpawn As Integer, secondpawn As Integer, values As Integer, automated As Boolean) As Boolean
If unloadthis = True Then
On Error Resume Next
End If
ReDim previousturn(UBound(gamepiece))
For x = 1 To UBound(gamepiece)
previousturn(x).boardposition = gamepiece(x).boardposition
previousturn(x).currentposition = gamepiece(x).currentposition
previousturn(x).leftss = sorrypieces(x).Left
previousturn(x).topss = sorrypieces(x).Top
previousturn(x).safetyposition = gamepiece(x).safetyposition
previousturn(x).status = gamepiece(x).status
Next


rollbacks = False

alreadyasked = False

continues = True
Dim newtext As Integer
On Error Resume Next
Close #1
On Error GoTo 0
If IsNumeric(Text1.Text) = False Then
Text1.Text = 7
End If
If automated = True Then
newtext = values
Else
newtext = Text1.Text
End If

Dim gameposition As positioning
Dim numbers As Integer

'MsgBox card(nextcard - 1).card & vbCrLf & card(nextcard - 1).description
'For z = 0 To List2.ListCount - 1
'If List2.Selected(z) = True Then
'y = z + 1
'Exit For
'End If
'Next
y = firstpawn
'MsgBox y

If card(nextcard - 1).card <> 1 And card(nextcard - 1).card <> 2 And card(nextcard - 1).card <> 13 And gamepiece(y).status = "Start" Then


MsgBox "Sorry, you must have a 1 or 2 to start or use the sorry card on someone"
taketurn = False

Exit Function
End If


numbers = 0
ReDim previous(1)
previous(1).gamepiece = y
previous(1).boardposition = gamepiece(y).boardposition
previous(1).safetyposition = gamepiece(y).safetyposition
previous(1).status = gamepiece(y).status
previous(1).color = gamepiece(y).color
previous(1).currentposition = gamepiece(y).currentposition
previous(1).starting = gamepiece(y).starting
previous(1).leftss = sorrypieces(y).Left
previous(1).topss = sorrypieces(y).Top

If card(nextcard - 1).card = 7 And newtext < 7 And gamepiece(y).status <> "Home" And gamepiece(y).status <> "Start" Then
If Check4.value = 1 Then


MsgBox "split so far"
End If

'new information
Dim splits As splitinformation
'MsgBox "Test"

splits = splitinformation(y, newtext, secondpawn)

If splits.cansplit = False Then
MsgBox "Sorry, you cannnot use this split"
taketurn = False
Exit Function
End If

'MsgBox splits.firstpawn & "   " & splits.firstposition
'MsgBox splits.secondpawn & "   " & splits.secondposition



ReDim Preserve previous(2)
previous(2).gamepiece = splits.secondpawn
previous(2).boardposition = gamepiece(splits.secondpawn).boardposition
previous(2).currentposition = gamepiece(splits.secondpawn).currentposition
previous(2).leftss = sorrypieces(splits.secondpawn).Left
previous(2).topss = sorrypieces(splits.secondpawn).Top
previous(2).status = gamepiece(splits.secondpawn).status


Dim thissplit() As positioning
thissplit = splitmove(splits)
For xx = 1 To 2
'MsgBox thissplit(xx).leftss & " for left " & xx & "  " & thissplit(xx).topss & " for top " & xx


If xx = 1 Then
sorrypieces(splits.firstpawn).Left = thissplit(1).leftss
sorrypieces(splits.firstpawn).Top = thissplit(1).topss
Else
sorrypieces(splits.secondpawn).Left = thissplit(2).leftss
sorrypieces(splits.secondpawn).Top = thissplit(2).topss
End If
Next




If sliderule(splits.firstpawn, True, True) = False And continues = True Then
sorrypieces(ngamepiece).Left = thissplit(1).leftss
sorrypieces(ngamepiece).Top = thissplit(1).topss
End If
If continues = False Then
taketurn = False
sorrypieces(splits.secondpawn).Left = previous(2).leftss
sorrypieces(splits.secondpawn).Top = previous(2).topss


Exit Function
End If

If continues = True And sliderule(splits.secondpawn, True, True) = False Then
sorrypieces(splits.secondpawn).Left = thissplit(2).leftss
sorrypieces(splits.secondpawn).Top = thissplit(2).topss
End If

'If sorrypieces(previous(1).gamepiece).Left = previous(1).leftss And sorrypieces(previous(1).gamepiece).Top = previous(1).topss And sorrypieces(previous(2).gamepiece).Left = previous(2).leftss And sorrypieces(previous(2).gamepiece).Top = previous(2).topss Then
If rollbacks = True Then
taketurn = False
Else
taketurn = True
End If




Exit Function




'put into array




End If


If card(nextcard - 1).card = 11 And Check1.value = 1 Then
'For z = 1 To UBound(gamepiece)
'If z <> y + 1 And gamepiece(z).color <> gamepiece(y).color And gamepiece(z).status = "Out" Then
'numbers = z
'Exit For
'End If
'Next
numbers = secondpawn

If numbers = 0 Then
MsgBox "No one to trade places"
taketurn = False

Exit Function
End If
opponent.boardposition = gamepiece(numbers).boardposition
opponent.safetyposition = gamepiece(numbers).safetyposition
opponent.status = gamepiece(numbers).status
opponent.color = gamepiece(numbers).color
opponent.gamepiece = y
opponent.currentposition = gamepiece(numbers).currentposition
opponent.starting = gamepiece(numbers).starting
opponent.leftss = sorrypieces(numbers).Left

opponent.topss = sorrypieces(numbers).Top

gameposition = tradeplaces(y, numbers)

If gameposition.noprocess = False Then
sorrypieces(y).Top = gameposition.topss
sorrypieces(y).Left = gameposition.leftss
taketurn = True
Else
taketurn = False
End If


Exit Function
End If



If card(nextcard - 1).card = 13 Then
If card(nextcard - 1).card = 13 And gamepiece(y).status <> "Start" Then
MsgBox "Sorry, you must use a man from start for the sorry"
taketurn = False

Exit Function
End If
numbers = secondpawn

'For z = 1 To UBound(gamepiece)
'If z <> y + 1 And gamepiece(z).color <> gamepiece(y).color And gamepiece(z).status = "Out" Then
'numbers = z
'Exit For
'End If
'Next

If numbers = 0 Then
MsgBox "No one to sorry"
taketurn = False
Exit Function

End If
End If


'y = 5
'MsgBox gamepiece(y).boardposition

z = numbers



If gamepiece(y).status <> "Home" Then
'MsgBox z

opponent.boardposition = gamepiece(z).boardposition
opponent.safetyposition = gamepiece(z).safetyposition
opponent.status = gamepiece(z).status
opponent.color = gamepiece(z).color
opponent.currentposition = gamepiece(z).currentposition
opponent.starting = gamepiece(z).starting
opponent.leftss = sorrypieces(z).Left
opponent.topss = sorrypieces(z).Top
opponent.gamepiece = z

gameposition = newposition(y, z, values)


If sliderule(y, True, True) = False Then


'MsgBox gamepiece(y).currentposition



'MsgBox y
'MsgBox gameposition.topss & "   " & gameposition.leftss



sorrypieces(y).Top = gameposition.topss
sorrypieces(y).Left = gameposition.leftss
'MsgBox gamepiece(y).boardposition
End If
'If sorrypieces(previous(1).gamepiece).Left = previous(1).leftss And sorrypieces(previous(1).gamepiece).Top = previous(1).topss Then
If rollbacks = True Then
taketurn = False
Else
taketurn = True
End If



Else
MsgBox "Sorry, you are already home"
taketurn = False

End If

End Function

Private Sub firstprocessmanuel()
'MsgBox "Test so far"
Dim x As Integer
manuelturn.card = card(nextcard - 1).card
manuelturn.firstpawn = 0
manuelturn.secondpawn = 0
manuelturn.opponent = 0
manuelturn.samecolor = 0
manuelturn.value = 0
z = 0

Dim latestpawn As Integer
For x = players(turnnum).startnum To players(turnnum).startnum + 3
If gamepiece(x).status = "Out" Then
z = z + 1



End If
Next


manuelturn.samecolor = z
'outss = z

z = 0




For x = 1 To UBound(gamepiece)
If gamepiece(x).status = "Out" And gamepiece(x).color <> players(turnnum).color Then
z = z + 1
latestpawn = x
End If
Next
manuelturn.opponent = z
If manuelturn.opponent = 1 And manuelturn.card = 11 Then
manuelturn.secondpawn = latestpawn
End If


If manuelturn.card = 4 Then
manuelturn.value = -4
End If
Dim allows As Boolean

For x = 1 To UBound(gamepiece)

allows = pawnallow(x, False)
If allows = True Then
sorrypieces(x).Enabled = True
Else
sorrypieces(x).Enabled = False
End If


Next
If manuelturn.opponent <> 0 And manuelturn.card = 11 Then
Check1.Enabled = True
End If
If manuelturn.card = 7 Then
Text1.Enabled = True
End If

Dim newlabels As String

todolabel.Visible = True
todo.Visible = True
newlabel = labelsay(1)
todo.Caption = newlabel

redos.Enabled = True


End Sub
Private Function labelsay(stages As Integer) As String
If pass.Enabled = True Then
newlabels = " or pass your turn"
Else
newlabels = ""
End If



If stages = 1 And manuelturn.card = 13 Then

labelsay = "Please choose a player to sorry" & newlabels
Exit Function
End If
If stages = 1 And card(nextcard - 1).card = 11 And Check1.Enabled = True Then
labelsay = "Please either move 11 spaces or trade places with an opponent" & newlabels
'MsgBox card(nextcard - 1).card & "   possible error"


Exit Function
End If
If stages = 1 And card(nextcard - 1).card = 7 And manuelturn.samecolor > 1 Then
labelsay = "Please either move 7 spaces or split your move between your pawns"
Exit Function
End If
If stages = 1 And card(nextcard - 1).card = 7 And manuelturn.samecolor = 1 Then
labelsay = "Please choose a pawn to move 7 spaces" & newlabels
Exit Function
End If

If stages = 1 And card(nextcard - 1).card = 10 Then
labelsay = "Please choose backwards 1 or forward 10 and pawn" & newlabels
Exit Function
End If

If stages = 1 And card(nextcard - 1).card = 1 And willmoveout.willmove = True And manuelturn.samecolor <> 0 Or stages = 1 And manuelturn.card = 2 And willmoveout.willmove = True And manuelturn.samecolor <> 0 Then
labelsay = "Please either move a pawn " & card(nextcard - 1).card & " spaces or take a pawn out of start"
Exit Function
End If

If stages = 1 And card(nextcard - 1).card = 1 And willmoveout.willmove = False And manuelturn.samecolor <> 0 Or stages = 1 And manuelturn.card = 2 And willmoveout.willmove = False And manuelturn.samecolor <> 0 Then
labelsay = "Please choose a pawn to move " & card(nextcard - 1).card & " spaces" & newlabels


Exit Function
End If
If stages = 1 And card(nextcard - 1).card = 4 Then
labelsay = "Please choose a pawn to move backwards 4" & newlabels
Exit Function
End If


If stages = 1 Then
labelsay = "Please choose a pawn to move " & manuelturn.card & " spaces" & newlabels
Exit Function
End If


If stages = 2 And card(nextcard - 1).card = 11 Then
labelsay = "Please choose the opponent to trade with"
Exit Function
End If
If stages = 2 And card(nextcard - 1).card = 7 Then
labelsay = "Please choose a pawn to split with"
Exit Function
End If



End Function

Private Function pawnallow(pawn As Integer, finalprocess As Boolean) As Boolean
pawnallow = False

Dim newleft As Integer



'If gamepiece(x).status = "Out" Or gamepiece(x).status = "Safety" Then


If gamepiece(pawn).status = "Safety" And gamepiece(pawn).safetyposition = 5 Then
newleft = 1
ElseIf gamepiece(pawn).status = "Safety" And gamepiece(pawn).safetyposition = 4 Then
newleft = 2
ElseIf gamepiece(pawn).status = "Safety" And gamepiece(pawn).safetyposition = 3 Then
newleft = 3
ElseIf gamepiece(pawn).status = "Safety" And gamepiece(pawn).safetyposition = 2 Then
newleft = 4
ElseIf gamepiece(pawn).status = "Safety" And gamepiece(pawn).safetyposition = 1 Then
newleft = 5
ElseIf gamepiece(pawn).status = "Out" And gamepiece(pawn).currentposition = 59 Then
newleft = 59 + 6
'MsgBox x

Else
newleft = 58 - gamepiece(pawn).currentposition + 6
End If


If finalprocess = True And manuelturn.firstpawn = pawn Then
pawnallow = False
Exit Function
End If


If card(nextcard - 1).card = 1 And willmoveout.pawn = pawn Or card(nextcard - 1).card = 2 And willmoveout.pawn = pawn Then
pawnallow = True
Exit Function
End If
If card(nextcard - 1).card = 13 And players(turnnum).color <> gamepiece(pawn).color And gamepiece(pawn).status = "Out" Then
pawnallow = True
Exit Function
End If
If card(nextcard - 1).card = 13 And players(turnnum).color = gamepiece(pawn).color Then
pawnallow = False
Exit Function
End If
If card(nextcard - 1).card = 11 And finalprocess = True And players(turnnum).color <> gamepiece(pawn).color And gamepiece(pawn).status = "Out" Then
pawnallow = True
Exit Function
End If
If card(nextcard - 1).card = 11 And finalprocess = False And players(turnnum).color = gamepiece(pawn).color And gamepiece(pawn).status = "Out" Then
pawnallow = True
Exit Function
End If
If card(nextcard - 1).card = 7 And players(turnnum).color = gamepiece(pawn).color And gamepiece(pawn).status <> "Home" And gamepiece(pawn).status <> "Start" Then
pawnallow = True
Exit Function
End If

If card(nextcard - 1).card = 10 And players(turnnum).color = gamepiece(pawn).color And gamepiece(pawn).status <> "Home" And gamepiece(pawn).status <> "Start" Then
pawnallow = True
Exit Function
End If
If card(nextcard - 1).card Then
newcard = -4
Else
newcard = card(nextcard - 1).card
End If


If newcard <= newleft And players(turnnum).color = gamepiece(pawn).color And gamepiece(pawn).status <> "Home" And gamepiece(pawn).status <> "Start" Then
pawnallow = True
Exit Function
End If



End Function
Private Sub disableall()
For x = 1 To UBound(gamepiece)
sorrypieces(x).Enabled = False
Next
redos.Enabled = False

End Sub

Private Sub Command9_Click()



'previousturn.boardposition = gamepiece(x).boardposition
'previousturn.currentposition = gamepiece(x).currentposition
'previousturn.leftss = sorrypieces(x).Left
'previousturn.topss = sorrypieces(x).topss
'previousturn.safetyposition = gamepiece(x).safetyposition
'previousturn.status = gamepiece(x).status
'Next
End Sub
Private Sub addcolumnss()
Dim newx As Integer
newx = ListView1.Width / 5
newx = newx - 100
ListView1.ColumnHeaders.Add , , "Pawn", newx
ListView1.ColumnHeaders.Add , , "B. P.", newx
ListView1.ColumnHeaders.Add , , "C. P.", newx
ListView1.ColumnHeaders.Add , , "Status", newx
ListView1.ColumnHeaders.Add , , "Safety", newx
ListView1.View = lvwReport

End Sub
Private Sub updatelists()

'Dim newx As Integer
'newx = ListView1.Width / 5
'newx = newx - 100
Do Until ListView1.ListItems.Count = 0
        ListView1.ListItems.Remove 1
    Loop
For x = 1 To UBound(gamepiece)
Set itmx = ListView1.ListItems.Add(, , gamepiece(x).color & "  " & x)
itmx.SubItems(1) = gamepiece(x).boardposition
itmx.SubItems(2) = gamepiece(x).currentposition
itmx.SubItems(3) = gamepiece(x).status
itmx.SubItems(4) = gamepiece(x).safetyposition
Next
ListView1.Refresh


End Sub

Private Sub Form_Load()
unloadthis = False

Picture1(2).Picture = LoadPicture(App.Path & "\sorry3.jpg")
Image2(2).Picture = LoadPicture(App.Path & "\sorry card2.jpg")


rollbacks = False
If Form2.Check2.value = 1 Then
testing = True
addcolumnss
ListView1.Visible = True
Text2.Visible = True
'rollback.Visible = True
'Text2.Enabled = True
Text3.Visible = True
Text3.Enabled = True
Label8.Visible = True
Label3.Visible = True
'Command3.Visible = True
'Command3.Enabled = True
'Command4.Visible = True
'Command4.Enabled = True
'Command5.Visible = True
'Command5.Enabled = True
'Command6.Visible = True
'Command6.Enabled = True
Check4.Enabled = True
Check4.Visible = True
Command8.Visible = True
Command8.Enabled = True
'Command1.Visible = True
'Command1.Enabled = True
Command7.Enabled = True
Command7.Visible = True
List2.Enabled = True
List2.Visible = True
Else
rollback.Visible = False
ListView1.Visible = False
Command7.Enabled = False
Command7.Visible = False
testing = False
Text2.Visible = False
Text2.Enabled = False
Text3.Visible = False
Text3.Enabled = False
Label8.Visible = False
Label3.Visible = False
Command3.Visible = False
Command3.Enabled = False
Command4.Visible = False
Command4.Enabled = False
Command5.Visible = False
Command5.Enabled = False
Command6.Visible = False
Command6.Enabled = False
Check4.Enabled = False
Check4.Visible = False
Command8.Visible = False
Command8.Enabled = False
Command1.Visible = False
Command1.Enabled = False
List2.Enabled = False
List2.Visible = False
End If
secondss = 2



'Set fso = New FileSystemObject
'Set t = fso.CreateTextFile(app.path\red safety.txt", True)
'Load sorrypieces(1)
'Set sorrypieces(1).Container = Picture1(2)
'sorrypieces(1).Picture = LoadPicture(app.path\red game piece.jpg")
'sorrypieces(1).Top = 0
'sorrypieces(1).Left = 0
'sorrypieces(1).Visible = True'
Command2.Visible = False
continues = True

'Exit Sub
ReDim players(0)

ReDim card(45)
shufflecards


y = 0
ReDim gamepiece(0)
w = 0
Dim x As Integer
For x = 0 To 3



If Form2.Check1(x).value = 1 Then

w = w + 1

'Open app.path\blue start.txt" For Input As #1
newpath = App.Path & "\" & Form2.Check1(x).Caption & " start.txt"
'MsgBox newpath
ReDim Preserve players(w)
players(w).playername = Form2.Text1(x + 1).Text
players(w).color = Form2.Check1(x).Caption
players(w).startnum = y + 1
'MsgBox w
'MsgBox y + 1

Open newpath For Input As #1


For z = 1 To 4
y = y + 1
'MsgBox Y
ReDim Preserve gamepiece(y)
List2.AddItem Form2.Check1(x).Caption & " " & z
gamepiece(y).color = Form2.Check1(x).Caption
gamepiece(y).status = "Start"
gamepiece(y).currentposition = 0
gamepiece(y).safetyposition = 0
gamepiece(y).starting = Form2.Check1(x).Tag
gamepiece(y).boardposition = 0


Line Input #1, c
searches = InStr(c, "   ")
leftss = Mid(c, 1, searches - 1)
topss = Mid(c, searches, Len(c) - searches + 1)
Load sorrypieces(y)
sorrypieces(y).Picture = LoadPicture(App.Path & "\" & Form2.Check1(x).Caption & " game piece.jpg")



Set sorrypieces(y).Container = Picture1(2)
sorrypieces(y).Top = topss
sorrypieces(y).Left = leftss
sorrypieces(y).Visible = True

'topss = Mid(c, searches + 2, Len(c) - searches - 3)
'MsgBox leftss
'MsgBox topss

'blue(x).Top = topss
'blue(x).Left = leftss
Next
Close #1
End If
Next
Open App.Path & "\sorry positions round.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, c
List1.AddItem c
Loop
Close #1
disableall

'MsgBox 16 + 16 + 14 + 14
'Set fso = New FileSystemObject
'Set t = fso.CreateTextFile("F:\Personal Information\Game Information\green safety.txt", True)

'Set t = fso.CreateTextFile(app.path positions.txt", True)
'Set Y = fso.CreateTextFile("F:\Personal Information\Game Information\blue home.txt", True)
'For X = 1 To 4
'Y.WriteLine blue(X).Left & "   " & blue(X).Top
'Next
'Y.Close
'Dim newsorrycard As cardinfo


For x = 1 To 45
card(x) = sorrycard(card(x).number)

'MsgBox card(x).number

'newsorrycard = sorrycard(x)
'MsgBox card(x).description & vbCrLf & card(x).card & vbCrLf & card(x).number
Next
nextcard = 1

turnnum = 0
newturn

'Unload Me
drawcards


'Set Y = fso.CreateTextFile("F:\Personal Information\Game Information\yellow home.txt", True)
'For X = 1 To 4
'Y.WriteLine yellow(X).Left & "   " & yellow(X).Top
'Next
'Y.Close
'y = 1



'Dim gameposition As positioning
'gameposition = newposition(y, 0)

'sorrypieces(y).Top = gameposition.topss
'sorrypieces(y).Left = gameposition.leftss

'sorrypieces(y).Top = topss
'sorrypieces(y).Left = leftss
'Set Y = fso.CreateTextFile("F:\Personal Information\Game Information\green home.txt", True)
'For X = 1 To 4
'Y.WriteLine green(X).Left & "   " & green(X).Top
'Next
'Y.Close


'Set Y = fso.CreateTextFile("F:\Personal Information\Game Information\red home.txt", True)
'For X = 1 To 4
'Y.WriteLine red(X).Left & "   " & red(X).Top
'Next
'Y.Close
'Unload Me
On Error Resume Next

List2.Selected(0) = True

Form2.Visible = False
Me.Show
If testing = False Then

status.Caption = "Setting up game"

'Pause (2)

processthismove


End If

'player.Caption = players(turnnum).playername
'color.Caption = players(turnnum).color

End Sub

Private Sub Form_Unload(Cancel As Integer)
't.Close
On Error Resume Next
'MsgBox "Test"
unloadthis = True

Unload Form2
'Unload Me

End Sub

Private Sub Image5_Click()
If nextcard > 45 Then
MsgBox "End of deck"
Else
drawcards



End If

End Sub
Sub drawcards()
If testing = True Then
updatelists
End If
'MsgBox "Test"


status.Caption = "Drawing next card"

If nextcard > 45 Then
status.Caption = "Shuffling cards"

ReDim card(45)
shufflecards

nextcard = 1
For x = 1 To 45
card(x) = sorrycard(card(x).number)

'MsgBox card(x).number

'newsorrycard = sorrycard(x)
'MsgBox card(x).description & vbCrLf & card(x).card & vbCrLf & card(x).number
Next
End If
status.Caption = "Drawing next card"


Image5.Picture = LoadPicture(App.Path & "\Sorry Card " & card(nextcard).card & ".jpg")


'MsgBox nextcard

Image5.ToolTipText = card(nextcard).description
'If card(nextcard).card = 13 Then
'Command2.Visible = True
'Command1.Visible = False
'y = 1

'check1 for 11
'Check1.Visible = False
'Label2.Visible = False



If card(nextcard).card = 11 Then

Check1.Visible = True
Label2.Visible = False
Text1.Visible = False
Option1.Visible = False
Option2.Visible = False
Check1.value = 0
ElseIf card(nextcard).card = 7 Then
Label2.Visible = True
Text1.Visible = True
Text1.Text = 7

Option1.Visible = False
Option2.Visible = False
Check1.Visible = False
ElseIf card(nextcard).card = 10 Then
Option1.Visible = True
Option2.Visible = True
Option1.value = True
Label2.Visible = False
Text1.Visible = False
Check1.Visible = False

Else
Option1.Visible = False
Option2.Visible = False
Label2.Visible = False
Text1.Visible = False
Check1.Visible = False
End If




'Else
'Command2.Visible = False
'Command1.Visible = True
'End If


nextcard = nextcard + 1
status.Caption = "Waiting for user input"
pass.Enabled = False

End Sub
'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox X
'MsgBox Y

Private Sub pass_Click()
newturn
drawcards
End Sub

'End Sub

Private Sub Picture1_Click(index As Integer)

End Sub

Private Sub Picture1_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Label1.Caption = ""
'sorrypieces(1).Left = x
'sorrypieces(1).Top = y

'blue(0).Left = X
'blue(0).Top = Y

'Image6(2).Left = X
'Image6(2).Top = Y

End Sub

Private Sub redos_Click()
firstprocessmanuel
Picture1(2).SetFocus

End Sub
Private Sub finalmanualprocess(index As Integer)
Dim winnings As Boolean

If card(nextcard - 1).card = 13 Then
'If Check4.value = 4 Then
'MsgBox willmoveout.pawn
'End If

manuelturn.firstpawn = willmoveout.pawn
manuelturn.secondpawn = index
manuelturn.value = 13
ElseIf card(nextcard - 1).card = 4 Then
manuelturn.firstpawn = index
manuelturn.value = -4
ElseIf Option2.value = True And card(nextcard - 1).card = 10 Then
manuelturn.value = -1
manuelturn.firstpawn = index
ElseIf gamepiece(index).status = "Start" And card(nextcard - 1).card < 3 Then
manuelturn.value = 0
manuelturn.firstpawn = index
ElseIf card(nextcard - 1).card <> 7 And card(nextcard - 1).card <> 11 Then
manuelturn.value = card(nextcard - 1).card
manuelturn.firstpawn = index
End If



Dim endturns As Boolean
'MsgBox manuelturn.firstpawn & "   " & manuelturn.value

endturns = taketurn(manuelturn.firstpawn, manuelturn.secondpawn, manuelturn.value, False)
If unloadthis = True Then
Exit Sub
End If

If endturns = True And card(nextcard - 1).card <> 2 Then
'Pause (secondss)

If winnings = True Then
Unload Me
Exit Sub
End If

newturn


drawcards


todo.Visible = False
todo.Caption = ""
todolabel.Visible = False
If testing = False Then
processthismove
End If



ElseIf endturns = True And card(nextcard - 1).card = 2 Then
'Pause (secondss)

winnings = haswon
If winnings = True Then
Unload Me
Exit Sub
End If



drawcards
todo.Visible = False
todo.Caption = ""
If testing = False Then
processthismove
End If

End If


End Sub

Private Sub rollback_Click()
On Error Resume Next

For x = 1 To UBound(gamepiece)
sorrypieces(x).Left = previousturn(x).leftss
sorrypieces(x).Top = previousturn(x).topss
gamepiece(x).boardposition = previousturn(x).boardposition
gamepiece(x).currentposition = previousturn(x).currentposition
gamepiece(x).safetyposition = previousturn(x).safetyposition
gamepiece(x).status = previousturn(x).status
Next
nextcard = nextcard - 2
drawcards

If card(nextcard - 1).card <> 2 Then

previousturns
'drawcards
End If

If Err.number <> 0 Then
MsgBox "Could not roll back because turn was not taken."

End If

rollback.Enabled = False

End Sub

Private Sub sorrypieces_Click(index As Integer)
If Check4.value = 1 Then
MsgBox index
End If

Dim x As Integer
If card(nextcard - 1).card < 7 Or card(nextcard - 1).card = 8 Or card(nextcard - 1).card = 10 Or card(nextcard - 1).card = 12 Or card(nextcard - 1).card = 13 Then
finalmanualprocess index
Exit Sub
End If

If card(nextcard - 1).card = 11 And Check1.value = 1 And Check1.Enabled = True And manuelturn.opponent = 1 Then

manuelturn.firstpawn = index
finalmanualprocess index
Exit Sub
End If

If card(nextcard - 1).card = 11 And Check1.value = 0 Then
'MsgBox manuelturn.secondpawn
manuelturn.firstpawn = index
manuelturn.value = 11

finalmanualprocess index

Exit Sub
End If

If card(nextcard - 1).card = 7 And Text1.Text >= 7 Then
manuelturn.firstpawn = index
manuelturn.value = 7
finalmanualprocess index
Exit Sub
End If

If card(nextcard - 1).card = 7 And IsNumeric(Text1.Text) = False Then
finalmanualprocess index
Exit Sub
End If

If card(nextcard - 1).card = 7 Then
z = 0
'Text1.Enabled = False
For x = players(turnnum).startnum To players(turnnum).startnum + 3
Dim latestpawn As Integer
If x <> index And gamepiece(x).status <> "Home" And gamepiece(x).status <> "Start" Then
latestpawn = x
z = z + 1

End If

Next
Text1.Enabled = False
If Check4.value = 1 Then
MsgBox z
End If

'If z = 0 Then
'MsgBox "Sorry, you cannot do this split because there is only one pawn out"
If z = 1 Then
manuelturn.firstpawn = index
manuelturn.secondpawn = latestpawn
manuelturn.value = Text1.Text
finalmanualprocess index
Exit Sub

End If
End If


If manuelturn.firstpawn <> 0 And manuelturn.card = 7 Then
manuelturn.secondpawn = index
finalmanualprocess index
Exit Sub
End If
If manuelturn.firstpawn <> 0 And manuelturn.card = 11 Then
manuelturn.secondpawn = index
finalmanualprocess index
Exit Sub
End If


If manuelturn.firstpawn = 0 And manuelturn.card = 7 Or manuelturn.card = 11 And manuelturn.firstpawn = 0 Then
Check1.Enabled = False
Text1.Enabled = False



For x = 1 To UBound(gamepiece)

allows = pawnallow(x, True)
If allows = True Then
sorrypieces(x).Enabled = True
Else
sorrypieces(x).Enabled = False
End If


Next

Dim newlabels As String

todolabel.Visible = True
todo.Visible = True
newlabel = labelsay(2)
todo.Caption = newlabel

redos.Enabled = True

If manuelturn.card = 7 Then
manuelturn.value = Text1.Text
manuelturn.firstpawn = index
ElseIf manuelturn.card = 11 Then
manuelturn.value = 11
manuelturn.firstpawn = index
End If
End If


'List2.Selected(Index - 1) = True

'MsgBox "This was enabled"






't.WriteLine sorrypieces(1).Left & "  " & sorrypieces(1).Top
'Label1.Caption = "Done"

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And IsNumeric(Text2.Text) = True Then
If Text2.Text > 13 Then
Exit Sub
End If
nextcard = nextcard - 1

card(nextcard).card = Text2.Text
card(nextcard).description = ""
drawcards



End If

End Sub
