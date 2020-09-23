VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Walking Guys"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label HimSay 
      BackStyle       =   0  'Transparent
      Caption         =   "HimSay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label MeSay 
      BackStyle       =   0  'Transparent
      Caption         =   "MeSay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image GuySkin 
      Height          =   480
      Index           =   4
      Left            =   9000
      Picture         =   "WalkingGuys.frx":0000
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image GuySkin 
      Height          =   480
      Index           =   3
      Left            =   9000
      Picture         =   "WalkingGuys.frx":040A
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image GuySkin 
      Height          =   480
      Index           =   2
      Left            =   9000
      Picture         =   "WalkingGuys.frx":07F4
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image GuySkin 
      Height          =   480
      Index           =   1
      Left            =   9000
      Picture         =   "WalkingGuys.frx":0BEA
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image GuySkin 
      Height          =   480
      Index           =   0
      Left            =   9000
      Picture         =   "WalkingGuys.frx":0FFA
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picbackground 
      Height          =   960
      Index           =   0
      Left            =   8760
      Picture         =   "WalkingGuys.frx":1427
      Top             =   4920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblhosting 
      BackColor       =   &H00FF80FF&
      Caption         =   "Hosting"
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
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
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblConnected 
      BackColor       =   &H0000FF00&
      Caption         =   "Connected"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.Image Guy 
      Height          =   480
      Index           =   2
      Left            =   6120
      Picture         =   "WalkingGuys.frx":3A69
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Guy 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "WalkingGuys.frx":3E73
      Top             =   3120
      Width           =   480
   End
   Begin VB.Menu mConnectHost 
      Caption         =   "Connect/Host"
      Begin VB.Menu mHost 
         Caption         =   "Host a game"
      End
      Begin VB.Menu mConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mgetIp 
         Caption         =   "Get External IP"
      End
   End
   Begin VB.Menu mOption 
      Caption         =   "Options"
      Begin VB.Menu mPort 
         Caption         =   "Choose Port"
      End
      Begin VB.Menu mDrawLines 
         Caption         =   "Draw Lines"
      End
      Begin VB.Menu mSay 
         Caption         =   "Say"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSkins 
         Caption         =   "Skins"
         Begin VB.Menu mSkin 
            Caption         =   "Warrior"
            Index           =   0
         End
         Begin VB.Menu mSkin 
            Caption         =   "Monk"
            Index           =   1
         End
         Begin VB.Menu mSkin 
            Caption         =   "Pirate"
            Index           =   2
         End
         Begin VB.Menu mSkin 
            Caption         =   "Witch"
            Index           =   3
         End
         Begin VB.Menu mSkin 
            Caption         =   "Hero"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SIMPLE MULTIPLAYER GAME
'Created by Mephisto

'Understanding of Cells
'This whole program is based on cells as you can notice. That means that i am creating
'squares on the form, where the guys can move. So i am limiting the positions they can be in
'This has its up and downs, down side is that you cannot see your guy nicely going from one
'position to another, he looks as though he is jumping there, appearing out of nowhere. But the
'good part about making cells is that you can monitor your guys, then later on in game you
'can add some objects in the form, stones, monsters, you get to know exact positions where they are
'and you dont have to use the RECT collision detection

'Also remember, winsock cant be perfect. If you have AOL you can lag. Connection, computer...
'some packets get lost, mixed up or even joined. We will discuss this later

'The next known bug is that when you start this application in Visual Basic and someone connects to you,
'then you try to say something, the text you send he wont see! BUt you will see the text he sends.
'I wasnt able to identify why this happens, but it is like that and i cant fix it
'So what you do when you want to use SAY feature appropriatly, create an EXE of this application
'and run that EXE, the SAY feature will be flawless when you run this as EXE...

'for FOR..NEXT loops
Dim i As Integer, r As Integer, c As Integer

Dim strr As String

'positions of both guys
Dim GuyXpos(2) As Integer
Dim GuyYpos(2) As Integer

'if lines option is selected
Dim lines As Boolean

'winsock stuff
Dim port As Single, hostIP As String

'winsock data arrival Dims, they will be needed later on
Dim Data As String
Dim Data2() As String
Dim data3() As String

'temporary variable for messagebox input yes/no
Dim temp As Integer

'for graphics loop in painting the picture with grass
Dim x As Integer, y As Integer

'to make sure there are no variables that we are using that are not dimmed
Option Explicit

Sub DrawLines()
'this procedure is called if the user selects the option to draw squares...

'this For..Next loop does horizontal lines
For i = 1 To 20
Form1.Line (i * 32, 0)-(i * 32, 480)
Next i

'this For..Next loop does vertical lines
For i = 1 To 15
Form1.Line (0, i * 32)-(640, i * 32)
Next i
'the form is 640 X 480 big

'this boolean holds the current state... false = no lines, true = draw lines
lines = True
End Sub

Private Sub Form_Click()
'if we click on form, both texts that characters say should be erased
HimSay.Visible = False
MeSay.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'this occurs when the user releases a key. We use this Select Case to find out which key
'the user released
Select Case KeyCode
'if its key up
Case vbKeyUp
    'then we first check if the Yposition of the player isnt 1, therefore he cant go more up!
    If GuyYpos(1) <> 1 Then
    'but if it is not, change his Y position one up
    GuyYpos(1) = GuyYpos(1) - 1
    End If
Case vbKeyRight
    'We check if the position of the player isnt already maximum of going right
    If GuyXpos(1) <> 20 Then
    GuyXpos(1) = GuyXpos(1) + 1
    End If
Case vbKeyDown
    'here we check if the player isnt at the bottom of the screen and he cant go any more down
    If GuyYpos(1) <> 15 Then
    GuyYpos(1) = GuyYpos(1) + 1
    End If
Case vbKeyLeft
    'if the user pressed left, but his position is already minimum left value, we cant go anymore left
    If GuyXpos(1) <> 1 Then
    GuyXpos(1) = GuyXpos(1) - 1
    End If
End Select

'here we move the actual guy to the updated X and Y positions
'we use the .Move command instead of setting .Left and .Top separatly which takes longer both to you
'and the computer to process it
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32

'finally, we need to send the new data over to the other player, we use On Error Resume Next,
'just to make sure the program doesnt crash if the user moves and he is not connected yet
On Error Resume Next
    
'This line can be difficult to understand so let me explain
'Programming with winsock is hard. If you are constantly sending a lot of packets at the same time,
'sometimes they get mixed up, or joined. Most of time though they are ok, but sometimes it doesnt
'work out. Lets have an example to illustrate this.
'lets say
'X = 5
'y = 15
'Now we need to send these over to the other side. We do this by joining these two numbers and putting
'a dummie between them, so we are sending one string in fact, but then later on on dataarrival section
'we extract each information by using the Split() function. Therefore we could do this:
'Winsock1.SendData GuyXpos(1) & "/" & GuyYpos(1)
'it seems like a good syntax from what i told you so far. But remember ! I said above that the packets
'sometimes join. So we are sending 5/15 very fast and eventually sometimes the other guy gets a string
'5/155/15. See the difference ? Two packets are joined and the Y coordinate becomes 155!
'Therefore we add "|" in the end. Its importance will be documented in dataarrival section.
    Winsock1.SendData GuyXpos(1) & "/" & GuyYpos(1) & "|"

End Sub

Private Sub Form_Load()
'to make sure all numbers generated are TRULY random
Randomize
'we set the default port
port = 5432
'we set the default position for both characters
GuyXpos(1) = 1
GuyYpos(1) = 1
GuyXpos(2) = 10
GuyYpos(2) = 10

'we move these characters to the default positions that were given above
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32

'this code is the code that creates the grass

    'first we make sure y is 0
    y = 0
    'then we cicle through the rows(r)
    For r = 1 To 14
        'everytime a row changes, we need to assign new X
        x = 0
        'we cycle through all columns
        For c = 1 To 20
            'we use this function to paint the form1 with the picture of grass
            Form1.PaintPicture picbackground(0).Picture, x, y
            'increase x by 32 (one square, or cell if you like)
           x = x + 32
        Next c
        'add 32 to Y
        y = y + 32
    Next r
    
    'insure graphic presistance
    Form1.Picture = Form1.Image

End Sub

Private Sub Form_Unload(Cancel As Integer)
'we want to close the winsock before we close the program
Winsock1.Close
End Sub

Private Sub mAbout_Click()
'if the user clicks About, display this message
MsgBox "This is a simple demonstration of how to use Winsock in order to play a multiplayer game through the internet with your friend." & vbCrLf & vbCrLf & "Thanks to Breezer for BETA testing!", , "About"
End Sub

Private Sub mConnect_Click()
On Error Resume Next
'If the user clicks Connect, we need to get the IP of the hosting computer first !
hostIP = InputBox("Enter the host's computer name or ip address:" & vbCrLf & "(Be careful not to include any unnecessary spaces etc or error message will be generated.")
'then we connect to this IP, on the default port
Winsock1.Connect hostIP, port
'we let the user know that he is connected
lblConnected.Visible = True

'we set up the starting default positions in a case that the user had moved before clicking connect
'notice, that in mHOST sub, we do the same thing, but we do it the other way,
'we let the GuyXpos(2) = 10, GuyYpos(2) = 10 and the GuyXpos(1) = 1 and GuyYpos(1) = 1
'This is a really hard part to explain, you really need to think about it. If you click connect,
'We set up your position 10, 10 and the host will be commanding the 1,1 guy
'In Host Sub, its vice versa, the Guy(1) is given the positions in 1,1, and the Guy2 in 10,10 !
GuyXpos(2) = 1
GuyYpos(2) = 1
GuyXpos(1) = 10
GuyYpos(1) = 10
'we move the pictures to updated positions
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32

End Sub

Private Sub mDrawLines_Click()
'If the user clicks DrawLines

'and Lines are already Drawen (This option of DrawLines works as toggle so we
'need to figure out the previous state)
If lines = True Then
    'clear the lines and set lines = false
    lines = False
    Form1.Cls
Else
' if the lines is False however,
    'we set lines = true and we call DrawLines procedure to make the actual lines
    lines = True
    DrawLines
End If
End Sub

Private Sub mDisconnect_Click()
'all we need to do here really is just to close Winsock1
Winsock1.Close
'and we let the user know that he is not hosting anymore and that he is not connected
lblhosting.Visible = False
lblConnected.Visible = False
End Sub

Private Sub mgetIp_Click()
'This Sub makes sure you get the external IP

'temporary Dims
Dim a As Integer, b As Integer
Dim strURL As String, strIP As String
'we open the www.whatismyip.com URL and read it whole into strURL, this can take some time
strURL = Inet1.OpenURL("http://www.whatismyip.com/")
'we find where the part before the IP is written
a = InStr(1, strURL, "<TITLE>Your ip is ")
'we find the part after the IP
b = InStr(1, strURL, " WhatIsMyIP.com</TITLE>")
' NOTE: the strURL doesnt have the actual Text you see in the browser in it, it has the HTML
'code of the site in it! You need to watch out for that

'And the IP itself is between these two!
strIP = Mid(strURL, a + 18, b - (a + 18))
'We let the user know what his IP is in form of message box
temp = MsgBox("Your IP is: " & strIP & vbCrLf & "Would you like to copy it to the clipboard ?", vbYesNo, "Your IP")
'if user clicked Yes and he wants to copy the IP into the clipboard
If temp = 6 Then
    'first clear the clipboard
    Clipboard.Clear
    'then assign new clipboard text - (our IP)
    Clipboard.SetText strIP
End If
End Sub

Private Sub mHost_Click()
'when we host, we make label lblhost visible to let the user know that he is hosting a game
lblhosting.Visible = True
'we assign local port
Winsock1.LocalPort = port
'and tell winsock to listen at the port
Winsock1.Listen
'we set the Xpositions and Ypositions. See the mConnect sub for more documentation!
'This is a crucial part of winsock understanding... compare this sub to the mConnect sub,
'Particulary the positions im assigning!
GuyXpos(1) = 1
GuyYpos(1) = 1
GuyXpos(2) = 10
GuyYpos(2) = 10

'we move the guys to the positions
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32
End Sub

Private Sub mPort_Click()
'We change the port
port = InputBox("Enter new port! I dont recommend lower port than 5000!!! default port = 5432")
End Sub

Private Sub mSay_Click()
'this is to make sure that if we are not connected and click SAY the program doesnt crash
On Error Resume Next
'what do you want to say?
strr = InputBox("What do you want to say?")
'if you didnt click Cancel or OK without typing anything
If strr <> "" Then
'we send the data with prefix 998 and again in the end we attach the | in case the packets are joined
Winsock1.SendData "998" & "/" & strr & "|"
'all crap :)
MeSay.Visible = True
MeSay.Left = Guy(1).Left
MeSay.Top = Guy(1).Top
MeSay.Caption = strr
End If
End Sub

Private Sub mSkin_Click(Index As Integer)

'if we are not connected
If lblConnected.Visible = False Then
MsgBox "NOTE: If you change the skin before you connect, the hosting computer will never recognize the change therefore you will be displayed as a normal hero."
End If

'pretty easy, we only read in a new image
Guy(1).Picture = GuySkin(Index).Picture

On Error Resume Next
'we send through the winsock these two > an identifier 999 and the index # of the skin
'in dataarrival sub, we check if the first data is 999 and if yes, we assign the skin with the # after
'to the guy(2) .... (see Data Arrival Sub for more documentation)
'notice the "|" in the end, we will discuss this in data arrival section
Winsock1.SendData "999" & "/" & Index & "|"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'if the other person is requesting connection, this Sub is executed

'if Winsock is already in use, we close it
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If

'we accept the user
Winsock1.Accept requestID
'we let the user know that we are connected to the other person and that he joined the game
lblConnected.Visible = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'This is the command to receive the data if there is any incoming
Winsock1.GetData Data, vbString, bytesTotal
'The data incoming is stored in the variable 'Data'

'This is the part that continues the explaination i gave you above, that the packets tend to join
'sometimes. I said we add "|" for this reason. Now lets take a look at two scenarios that can occur.
'1
'Lets say we send 5/15|
'In this procedure we look if the packets are joined. If they were joined they would look like this:
'5/15|5/15
'So we test. If the lenght of the data received is more then 6 and the packet joining has occured
If Len(Data) > 6 Then
'first we use a temporary variable to store the first part. We extract the 5/15 out of the 5/15|5/15
data3 = Split(Data, "|")
'then from the first part we extract both values Data2(0) becomes the X (5) and Data2(1) becomes Y (15)
Data2 = Split(data3(0), "/")
Else
'but if the packets joining hasnt occured and the Data arrived is 5/15| then
'we first split these terms by /, thus giving us 5 and 15|
Data2 = Split(Data, "/")
'I hope you notice that the Y coordinate is 15| instead of 15. So we need to cut the last character
'so now the Data2(0) is the X (5) and Data2(1) is the Y coordinate (15)
Data2(1) = Left$(Data2(1), Len(Data2(1)) - 1)
End If

'we print both
Label1.Caption = "Xpos = " & Data2(0) & "  .. Ypos = " & Data2(1)

'if the data2(0) is 999 and the skin change has occured
If CInt(Data2(0)) = 999 Then
'change the skin
Guy(2).Picture = GuySkin(CInt(Data2(1)))
'but if the prefix is 998 (for SAY message)
ElseIf CInt(Data2(0)) = 998 Then
'we do all the crap, display the label, put in text, move it appropriatly
HimSay.Visible = True
HimSay.Left = Guy(2).Left
HimSay.Top = Guy(2).Top
HimSay.Caption = Data2(1)
Else
'update the position of second player
Guy(2).Move CInt(Data2(0)) * 32 - 32, CInt(Data2(1)) * 32 - 32

End If
End Sub



