VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Morse Code Machine ------- Sherif Rofael."
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3000
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encode Morse"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Morse Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer m 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
m.FileName = ""
Kill App.Path & "\sample2.mid"
Dim AddDot() As Byte
Dim AddDash() As Byte
Dim Header() As Byte
Dim empty1() As Byte
ReDim empty1(1 To 3000) As Byte


Open App.Path & "\main1.dat" For Binary Access Read As #4
MY_Filesize = LOF(4)
ReDim Header(1 To MY_Filesize) As Byte
Get #4, , Header()
Close #4


Open App.Path & "\e.dat" For Binary Access Read As #1
MY_Filesize = LOF(1)
ReDim AddDot(1 To MY_Filesize) As Byte
Get #1, , AddDot()
Close #1

Open App.Path & "\t.dat" For Binary Access Read As #2
MY_Filesize = LOF(2)
ReDim AddDash(1 To MY_Filesize) As Byte
Get #2, , AddDash()
Close #2


Open App.Path & "\sample2.mid" For Binary Access Write As #3
'For i = 1 To 3

Put #3, , Header()
For i = 1 To Len(Text2.Text)
morsesymbol = Mid(Text2.Text, i, 1)
Select Case morsesymbol
Case "."
Put #3, , AddDot()
Case "-"
Put #3, , AddDash()
Case " "
Put #3, , empty1()
End Select
Next i
Close #3

End Sub

Function EncodeMorse(aCharachter)
aCharachter = LCase(aCharachter)
Select Case aCharachter
    Case " "
        EncodeCharachter = " // "
    Case "a"
        EncodeCharachter = ".-"
    Case "b"
        EncodeCharachter = "-..."
    Case "c"
        EncodeCharachter = "-.-."
    Case "d"
        EncodeCharachter = "-.."
    Case "e"
        EncodeCharachter = "."
    Case "f"
        EncodeCharachter = "..-."
    Case "g"
        EncodeCharachter = "--."
    Case "h"
        EncodeCharachter = "...."
    Case "i"
        EncodeCharachter = ".."
    Case "j"
        EncodeCharachter = ".---"
    Case "k"
        EncodeCharachter = "-.-"
    Case "l"
        EncodeCharachter = ".-.."
    Case "m"
        EncodeCharachter = "--"
    Case "n"
        EncodeCharachter = "-."
    Case "o"
        EncodeCharachter = "---"
    Case "p"
        EncodeCharachter = ".--."
    Case "q"
        EncodeCharachter = "--.-"
    Case "r"
        EncodeCharachter = ".-."
    Case "s"
        EncodeCharachter = "..."
    Case "t"
        EncodeCharachter = "-"
    Case "u"
        EncodeCharachter = "..-"
    Case "v"
        EncodeCharachter = "...-"
    Case "w"
        EncodeCharachter = ".--"
    Case "x"
        EncodeCharachter = "-..-"
    Case "y"
        EncodeCharachter = "-.--"
    Case "z"
        EncodeCharachter = "--.."
    Case "1"
        EncodeCharachter = ".----"
    Case "2"
        EncodeCharachter = "..---"
    Case "3"
        EncodeCharachter = "...--"
    Case "4"
        EncodeCharachter = "....-"
    Case "5"
        EncodeCharachter = "....."
    Case "6"
        EncodeCharachter = "-...."
    Case "7"
        EncodeCharachter = "--..."
    Case "8"
        EncodeCharachter = "---.."
    Case "9"
        EncodeCharachter = "----."
    Case "0"
        EncodeCharachter = "-----"
    Case "."
        EncodeCharachter = ".-.-.-"
    Case "?"
        EncodeCharachter = "..--.."
    Case ","
        EncodeCharachter = "--..--"
    Case "'"
        EncodeCharachter = ".----."
    Case Else
        EncodeCharachter = (aCharachter)
End Select
EncodeMorse = EncodeCharachter
End Function

Private Sub Command2_Click()
Text2.Text = ""
For i = 1 To Len(Text1.Text)
mm11 = Mid(Text1.Text, i, 1)

If mm11 <> Chr(10) And mm11 <> Chr(13) Then
Text2.Text = Text2.Text & EncodeMorse(mm11) & " "
Else
Text2.Text = Text2.Text & vbCrLf
i = i + 1
End If
Next i
Call Command1_Click
m.FileName = App.Path & "\sample2.mid"
m.AutoStart = False
m.Width = 1935
End Sub

