VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16752
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   16752
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLight 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Light Mode"
      Height          =   372
      Left            =   6600
      TabIndex        =   45
      Top             =   1140
      Width           =   1992
   End
   Begin VB.CommandButton cmdDark 
      Caption         =   "Dark Mode"
      Height          =   492
      Left            =   6600
      TabIndex        =   44
      Top             =   540
      Width           =   1992
   End
   Begin VB.CommandButton cmdlevel3 
      Caption         =   "level 3 "
      Height          =   492
      Left            =   6780
      TabIndex        =   42
      Top             =   6540
      Width           =   2532
   End
   Begin VB.CommandButton cmdLevel2 
      Caption         =   "level 2"
      Height          =   552
      Left            =   6780
      TabIndex        =   41
      Top             =   5880
      Width           =   2592
   End
   Begin VB.CommandButton cmd30 
      Caption         =   "30 SEC CHALLENGE"
      Height          =   672
      Left            =   10200
      TabIndex        =   40
      Top             =   5700
      Width           =   1692
   End
   Begin VB.CommandButton cmdChallenge 
      Caption         =   "1 MIN CHALLENGE"
      Height          =   672
      Left            =   10200
      TabIndex        =   39
      Top             =   6540
      Width           =   1752
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "EXIT"
      Height          =   732
      Left            =   14340
      TabIndex        =   37
      Top             =   4800
      Width           =   1452
   End
   Begin VB.Timer Timer2 
      Left            =   540
      Top             =   7620
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   60
      Top             =   7680
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H000000FF&
      Caption         =   " UNWRAP THE SUSPENCE. AM I RIGHT???"
      Height          =   792
      Left            =   14340
      TabIndex        =   34
      Top             =   3720
      Width           =   1452
   End
   Begin VB.ListBox lstLetters 
      Height          =   1968
      Left            =   11100
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.ListBox lstdash 
      Height          =   1968
      Left            =   12840
      TabIndex        =   29
      Top             =   780
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.ListBox lstWords 
      Height          =   2352
      Left            =   14940
      TabIndex        =   28
      Top             =   660
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton CmdPlay 
      BackColor       =   &H000000FF&
      Caption         =   "PLAY"
      Height          =   792
      Left            =   14340
      TabIndex        =   26
      Top             =   5700
      Width           =   1512
   End
   Begin VB.Label oh 
      Height          =   372
      Left            =   7140
      TabIndex        =   48
      Top             =   2580
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label chal 
      BackStyle       =   0  'Transparent
      Caption         =   "Challenegs"
      Height          =   312
      Left            =   10200
      TabIndex        =   47
      Top             =   5340
      Width           =   1272
   End
   Begin VB.Label ko 
      BackStyle       =   0  'Transparent
      Caption         =   "Levels"
      Height          =   312
      Left            =   6780
      TabIndex        =   46
      Top             =   5460
      Width           =   1212
   End
   Begin VB.Image imgDove 
      Height          =   1872
      Left            =   4680
      Picture         =   "frmUserPlay.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   1632
   End
   Begin VB.Image imgHP 
      Height          =   1692
      Left            =   9120
      Picture         =   "frmUserPlay.frx":12DA6
      Stretch         =   -1  'True
      Top             =   540
      Visible         =   0   'False
      Width           =   1752
   End
   Begin VB.Label t 
      BackColor       =   &H0000FF00&
      Caption         =   "Themes"
      Height          =   372
      Left            =   6660
      TabIndex        =   43
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label hat 
      BackColor       =   &H0000FF00&
      Caption         =   "TIMER"
      Height          =   252
      Left            =   4560
      TabIndex        =   38
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label CountDown 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4500
      TabIndex        =   36
      Top             =   3720
      Width           =   996
   End
   Begin VB.Label lblHangman 
      Caption         =   "HANGMAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   120
      TabIndex        =   35
      Top             =   6180
      Width           =   4332
   End
   Begin VB.Label lblWIN 
      BackColor       =   &H000000FF&
      Caption         =   "OH MY GOD YES! YOU WIN"
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   6600
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   2112
   End
   Begin VB.Label lblLOST 
      BackColor       =   &H000000FF&
      Caption         =   " OH NO YOU LOST"
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   6600
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   2232
   End
   Begin VB.Image imglegright 
      Height          =   1212
      Left            =   1800
      Picture         =   "frmUserPlay.frx":18B6B
      Stretch         =   -1  'True
      Top             =   3060
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Image imglegLeft 
      Height          =   1272
      Left            =   1020
      Picture         =   "frmUserPlay.frx":1C4A8
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   792
   End
   Begin VB.Image imghandright 
      Height          =   1572
      Left            =   2100
      Picture         =   "frmUserPlay.frx":1FB9B
      Stretch         =   -1  'True
      Top             =   1740
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Image imghandLeft 
      Height          =   1092
      Left            =   660
      Picture         =   "frmUserPlay.frx":320DB
      Stretch         =   -1  'True
      Top             =   2100
      Visible         =   0   'False
      Width           =   792
   End
   Begin VB.Image imgStomach 
      Height          =   1032
      Left            =   1200
      Picture         =   "frmUserPlay.frx":33A26
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.Image imgneckk 
      Height          =   732
      Left            =   1320
      Picture         =   "frmUserPlay.frx":3D68C
      Stretch         =   -1  'True
      Top             =   1740
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblword 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   9300
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.Image imgHead 
      Height          =   912
      Left            =   1440
      Picture         =   "frmUserPlay.frx":3E4BA
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   1620
      X2              =   1620
      Y1              =   720
      Y2              =   1140
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   540
      X2              =   1620
      Y1              =   660
      Y2              =   720
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   600
      X2              =   540
      Y1              =   4500
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   660
      X2              =   2400
      Y1              =   4440
      Y2              =   4500
   End
   Begin VB.Label lblwsf 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   9240
      TabIndex        =   27
      Top             =   3660
      Width           =   4572
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Z"
      Height          =   492
      Index           =   25
      Left            =   4320
      TabIndex        =   25
      Top             =   5400
      Width           =   732
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      Height          =   432
      Index           =   24
      Left            =   3660
      TabIndex        =   24
      Top             =   5460
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      Height          =   432
      Index           =   23
      Left            =   3000
      TabIndex        =   23
      Top             =   5460
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      Height          =   492
      Index           =   22
      Left            =   2340
      TabIndex        =   22
      Top             =   5460
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      Height          =   492
      Index           =   21
      Left            =   1740
      TabIndex        =   21
      Top             =   5460
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      Height          =   492
      Index           =   20
      Left            =   1080
      TabIndex        =   20
      Top             =   5460
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      Height          =   432
      Index           =   19
      Left            =   480
      TabIndex        =   19
      Top             =   5460
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      Height          =   432
      Index           =   18
      Left            =   11280
      TabIndex        =   18
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      Height          =   432
      Index           =   17
      Left            =   10800
      TabIndex        =   17
      Top             =   4860
      Width           =   432
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Q"
      Height          =   432
      Index           =   16
      Left            =   10260
      TabIndex        =   16
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      Height          =   432
      Index           =   15
      Left            =   9780
      TabIndex        =   15
      Top             =   4860
      Width           =   432
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      Height          =   432
      Index           =   14
      Left            =   9240
      TabIndex        =   14
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      Height          =   432
      Index           =   13
      Left            =   8700
      TabIndex        =   13
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      Height          =   432
      Index           =   12
      Left            =   8100
      TabIndex        =   12
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
      Height          =   492
      Index           =   11
      Left            =   7500
      TabIndex        =   11
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      Height          =   492
      Index           =   10
      Left            =   6960
      TabIndex        =   10
      Top             =   4860
      Width           =   492
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "J"
      Height          =   492
      Index           =   9
      Left            =   6360
      TabIndex        =   9
      Top             =   4860
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      Height          =   492
      Index           =   8
      Left            =   5700
      TabIndex        =   8
      Top             =   4860
      Width           =   612
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      Height          =   492
      Index           =   7
      Left            =   5040
      TabIndex        =   7
      Top             =   4860
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      Height          =   492
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   4860
      Width           =   672
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      Height          =   492
      Index           =   5
      Left            =   3660
      TabIndex        =   5
      Top             =   4860
      Width           =   612
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E"
      Height          =   492
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   4860
      Width           =   612
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      Height          =   492
      Index           =   3
      Left            =   2340
      TabIndex        =   3
      Top             =   4860
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      Height          =   492
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Top             =   4860
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      Height          =   492
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   4860
      Width           =   552
   End
   Begin VB.Label lblLetts 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      Height          =   492
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   4860
      Width           =   552
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lives As Integer ' count the number of tries?
'Dim wsf As Integer ' word so far is will have letters plus dashes.
Dim wsf As String
Dim arrletters(22) As String ' we know that the maximum letters in the word is 22, so we cr
Dim i As Long
Dim strLine As String ' read each line
Dim lwl As Integer ' ??
Dim numwords As Long
Dim arrwords(60000) As String ' nuber of words counter from dictionary.text
Dim big As String ' biggest word
Dim small As String ' smallest word
Dim word As String
Dim arrdashes(22) As String ' parallel to arrletters
Dim flag As String
'Dim flag As Integer
Dim lp As String
Dim pa As String
'Dim p As Integer
Dim index As Integer
Dim rnum As String
Dim nHeight As Integer
Dim n As Integer

Private Sub cmd30_Click()
Timer2.Interval = 1000
     CountDown.Caption = "30"
End Sub

Private Sub cmdChallenge_Click()
Timer2.Interval = 1000
     CountDown.Caption = "60"
End Sub

Private Sub cmdCheck_Click()
If wsf = word Then
    lblWIN.Visible = True
Else
    
    lblLOST.Visible = True
    End If
    
End Sub

Private Sub cmdDark_Click()
Form1.BackColor = &H80&

lblword.BackColor = &H80&

hat.BackColor = &H80&
imgHP.Visible = True
imgDove.Visible = False


End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLevel1_Click()
 
End Sub

Private Sub cmdLevel2_Click()
imgHead.Visible = True
imgneckk.Visible = True
imgStomach.Visible = True
lives = 3
End Sub

Private Sub cmdlevel3_Click()
imgHead.Visible = True
imgneckk.Visible = True
imgStomach.Visible = True
imghandLeft.Visible = True

lives = 4
End Sub

Private Sub cmdLight_Click()
Form1.BackColor = &HFFFFFF

lblword.BackColor = &HFFFFFF

hat.BackColor = &HFFFFFF
imgDove.Visible = True
imgHP.Visible = False

End Sub

Private Sub CmdPlay_Click()

'wsf = wsf + "-"
'wsf = wsf + arrdashes(i)
'For i = 1 To Len(word)
'lstWords = wsf + arrdashes(i)
'Next i
'hrfuehgguergg wyghefgbhwegfj ^^

'ws

Randomize

rnum = Int(Rnd * numwords) + 1
word = arrwords(rnum)
lives = 0
lblword.Caption = word
'End Sub

For i = 1 To Len(word)
    arrletters(i) = Mid(word, i, 1)
    lstLetters.AddItem (arrletters(i))
    arrdashes(i) = "-"
    lstdash.AddItem (arrdashes(i))
    
    
    flag = 0
Next i

' LETTER A OR C - NOW A IS IN SCOOL OR C IS IN SCHOOL SHLD BE BEHIND THE CODE. IS A = S

'Dim i As Integer
' For i = 1 To Len(word)
 '   arrletters(i) = Mid(word, i, 1)
  '  arrdashes(i) = ""
   ' flag = 0
 'Next i
    For i = 1 To Len(word)
     If lp = arrletters(i) Then
     flag = 1
     arrdashes(i) = lp
     End If
    wsf = wsf + arrdashes(i)
    
    Next i

'lblwsf.Caption = (arrdashes(Len(word)))
 lblwsf = wsf
 
 Timer2.Interval = 1000
     CountDown.Caption = "100"

End Sub
    
 



Private Sub form_load() ' ARREAY OF WORDS . PLAY = PARALLEL ARRATS AND DASHES

'opening file

Open "C:\Users\twish\Documents\problem solving 102\Labs\dictionary.txt" For Input As #1

' do while
 i = 1
 Do While Not EOF(1)
    Input #1, strLine
    strLine = Trim(strLine)
    If Len(strLine) <> 0 Then
    arrwords(i) = strLine
    lstWords.AddItem (arrwords(i))
    i = i + 1
    End If


 Loop

numwords = i - 1
'lblNumWords.Caption = numwords
' lblNumWords.Caption = numwords
'numwords = Val((arrwords(i)) - 1)
'lblNumWords.Caption = numwords
    
'  For i = 1 To maxline
 ' Line Input #1, strLine
  'strLine = Trim(strLine)
  'If Len(strLine) <> 0 Then
'  lstWords.AddItem (strLine)
 ' Next i
  '  ' close
    
Close #1
 ' End If
'; Loop
'numwords = lblNumWords.Caption
'lblNumWords.Caption = numwords


'lblNumWords = UBound(arrwords)
'(lstWords.ListCount - 1)


' ec
' For i = 0 To UBound(arrwords)
'For i = 0 To numwords
 'big = arrwords(i)
  '  If Len(arrwords(i)) > Len(big) Then
 '   big = arrwords(i)
  '  End If
   ' i = i + 1



' Next i
'lblword.Caption = big
'lbllength = Len(big)
'lbllength = lwl

''''OMGGG DJIEHO
' populate array
'a = Len(big)
'Dim arrLetter(40) As String
'For i = 0 To Len(word)
 '   arrletters(i) = Mid(word, i, 1)
  '  arrdashes(i) = "-"
'Next
'CONTROL ARRAY


'arrLetter(i) = big
'Mid(big,0,a)
'lstdash.AddItem (arrletters(i))

'Next i
'' end omg


' RANDOMLY GENERATIG WORD
'Randomize
'rnum = Int(Rnd * numwords) + 1
'word = arrwords(rnum)
 '   lblword.Caption = word
  '  wsf = ""
   ' flag = 0
'For i = 1 To lenword

'Timer2.Interval = 1000
     'CountDown.Caption = "100"


End Sub
Private Sub Timer2_Timer()
     If CountDown.Caption = 0 Then
          Timer2.Enabled = False
          MsgBox ("TIME IS UP up, You Lose")
          lblLOST.Visible = True
     Else
          CountDown.Caption = CountDown.Caption - 1
     End If
End Sub

Private Sub lblLetts_Click(index As Integer)

wsf = ""
flag = 0
lp = Chr(index + 97)
lblLetts(index).Enabled = False
For i = 1 To Len(word)
    If lp = arrletters(i) Then
        flag = 1
        arrdashes(i) = lp
    End If
    wsf = wsf + arrdashes(i)
Next i
lblwsf = wsf
If flag = 0 Then
    lives = lives + 1
    If lives = 1 Then
    imgHead.Visible = True
    End If
    If lives = 2 Then
    imgneckk.Visible = True
    End If
    If lives = 3 Then
    imgStomach.Visible = True
    End If
    If lives = 4 Then
    imghandLeft.Visible = True
    End If
    If lives = 5 Then
    imghandright.Visible = True
    End If
    If lives = 6 Then
    imglegLeft.Visible = True
    End If
    If lives = 7 Then
    imglegright.Visible = True
    End If
    If lives > 6 Then
    lblLOST.Visible = True
    Timer2.Enabled = False
    lblLetts(0).Visible = False
    lblLetts(1).Visible = False
    lblLetts(2).Visible = False
    lblLetts(3).Visible = False
    lblLetts(4).Visible = False
    lblLetts(5).Visible = False
    lblLetts(6).Visible = False
    lblLetts(7).Visible = False
    lblLetts(8).Visible = False
    lblLetts(9).Visible = False
    lblLetts(10).Visible = False
    lblLetts(11).Visible = False
    lblLetts(12).Visible = False
    lblLetts(13).Visible = False
    lblLetts(14).Visible = False
    lblLetts(15).Visible = False
    lblLetts(16).Visible = False
    lblLetts(17).Visible = False
    lblLetts(18).Visible = False
    lblLetts(19).Visible = False
    lblLetts(20).Visible = False
    lblLetts(21).Visible = False
    lblLetts(22).Visible = False
    lblLetts(23).Visible = False
    lblLetts(24).Visible = False
    lblLetts(25).Visible = False
    oh.Caption = wsf
    lblword.Visible = True
    
    End If
    'If lives > 6 And lblLOST.Visible = False Then
    'If arrletters(word) = a Then
    'If wsf = word Then
    'lblWIN.Visible = True
    'End If
    
    
End If

End Sub

Private Sub lblLett_click(index As Integer)

'lp = Chr(index + 97)
'lblLetts(index).Enabled = False

'flags needed ? 2 ( img) + 1(fail) = 3

'For i = 1 To Len(word)
'flag = 0

'For i = 1 To Len(word)
 'flag = 0
'For i = 1 To Len(word)
'flag = 0
'    If lp = arrletters(i) Then
 '      flag = 1
 '   Else
'        flag = 0
'    End If
'Next i
        
'Next i

        
        
        
            
            
            
        

End Sub

'opening file

'Open "C:\Users\twish\Documents\problem solving 102\Labs\dictionary.txt" For Input As #1

' do while
 'i = 1
 'Do While Not EOF(1)
  '  Input #1, strLine
   ' strLine = Trim(strLine)
    'If Len(strLine) <> 0 Then
 '   arrwords(i) = strLine
  '  lstWords.AddItem (arrwords(i))
   ' i = i + 1
   ' End If


 'Loop

'numwords = i - 1
'lblNumWords.Caption = numwords
' lblNumWords.Caption = numwords
'numwords = Val((arrwords(i)) - 1)
'lblNumWords.Caption = numwords
    
'  For i = 1 To maxline
 ' Line Input #1, strLine
  'strLine = Trim(strLine)
  'If Len(strLine) <> 0 Then
'  lstWords.AddItem (strLine)
 ' Next i
  '  ' close
    
'Close #1
 ' End If
'; Loop
'numwords = lblNumWords.Caption
'lblNumWords.Caption = numwords


'lblNumWords = UBound(arrwords)
'(lstWords.ListCount - 1)


' ec
' For i = 0 To UBound(arrwords)
'For i = 0 To numwords
' big = arrwords(i)
'    If Len(arrwords(i)) > Len(big) Then
 '   big = arrwords(i)
  '  End If
   ' i = i + 1



 'Next i
'lblword.Caption = big
'lbllength = Len(big)
'lbllength = lwl



' populate array
'a = Len(big)

'For i = 0 To Len(big)
'Dim arrLetter(40) As String
'arrLetter(i) = big
'mid(big,0,a)
'lst2.AddItem (arrLetter(i))

'Next i






Private Sub lblwsf_Click()

End Sub

Private Sub Timer_Timer()


' _Now is a function that returns the system date and time in a single value

' _Second gives a value between 0-59 equal to number of seconds of current time

n = Second(Now) Mod 10

 If n = 0 Then
 nHeight = 8.25
 lblHangman.BackColor = &H8000000D
 
 lblHangman.ForeColor = &H8000000D

lblHangman.Move Left

 '
   'OR
   'Label1.Left = ORIGINAL_LEFT_POSITION
 
 ElseIf n = 1 Or n = 9 Then
  nHeight = 9.75
  lblHangman.BackColor = &H80FF80
  lblHangman.ForeColor = &HC0&
  
  lblHangman.Left = lblHangman.Left + 190

   
  
  
  
  ElseIf n = 2 Or n = 8 Then
  nHeight = 12
  lblHangman.BackColor = &HFFC0FF
  lblHangman.ForeColor = &H0&
   lblHangman.Left = lblHangman.Left + 40
  
  ElseIf n = 3 Or n = 7 Then
  nHeight = 13.5
   lblHangman.BackColor = &HC00000
  lblHangman.ForeColor = &HFFFFFF
  
   lblHangman.Left = lblHangman.Left - 190
  
  ElseIf n = 4 Or n = 6 Then
  nHeight = 18
   lblHangman.BackColor = &H808000
  lblHangman.ForeColor = &HC00000
  
lblHangman.Left = lblHangman.Left - 60
  
  
  Else 'n= 5
   nHeight = 24
   
   End If
  
  
 
lblHangman.FontSize = nHeight
'lblHangman.Caption = Time$ ' as before

 


'lblHangman = Time

End Sub
