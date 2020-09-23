VERSION 5.00
Begin VB.Form frmSkin 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2700
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicLoadButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2330
      ScaleHeight     =   330
      ScaleWidth      =   1050
      TabIndex        =   8
      Top             =   2200
      Width           =   1080
   End
   Begin VB.PictureBox PicExitButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3480
      ScaleHeight     =   330
      ScaleWidth      =   1050
      TabIndex        =   5
      Top             =   2200
      Width           =   1080
   End
   Begin VB.PictureBox PicExit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4400
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   20
      Width           =   255
   End
   Begin VB.PictureBox PicMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4150
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   20
      Width           =   255
   End
   Begin VB.PictureBox PicTitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -10
      ScaleHeight     =   300
      ScaleWidth      =   4680
      TabIndex        =   2
      Top             =   -10
      Width           =   4710
   End
   Begin VB.PictureBox PicBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2420
      Left            =   -10
      ScaleHeight     =   2385
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   300
      Width           =   4710
      Begin VB.ListBox SkinLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   270
         ItemData        =   "frmSkin.frx":08CA
         Left            =   720
         List            =   "frmSkin.frx":08EC
         TabIndex        =   9
         Top             =   1965
         Width           =   615
      End
      Begin VB.Label EmailMSG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "     Please e-mail me at: TheLeadX@aol.com       for your Questions and/or comments."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label DisplayMSG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is an example on how to use BitBlt to skin your vb programs. This is just an example so don't laugh at my cheezy skins! :)"
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4470
      End
   End
   Begin VB.PictureBox PicSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   720
      ScaleHeight     =   0
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   0
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'Mike's Skin Example v1.0  |4-15-2000|
'**************************************

'Thanks for downloading my Skin Example v1.0,
'There are probobly a couple of reasons why you did:
'Either you were interested in what this, you
'Wanted to know how to skin your programs,
'Or you just wanted to learn how to use BitBlt.

'Is this example I made, it's a decent one on
'How to use BitBlt to Skin your apps. The skins
'Aint that great but the main thing to focus on is
'The coding and how it cut's sections of a Bitmap(bmp)
'Using BitBlt and then pastes it into a picturebox.
'When I made this I didn't know much about BitBlt either
'But I got pretty good after making this. So I recommend
'You try something like this from scratch and use my coding
'As a reference...It took me about an hour or two to make this
'Example(10 Skins, Commenting, Perfection..heh) so If I can do this
'And understand it then you can to. I commented "EVERY" line in this
'Example..Declarations to everything below. So enjoy and I hope you
'Learn something as I have.

'-Mike



'<START> of coding

'Makes sure everything is defined:
    Option Explicit
    
'Declarations:
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 'What this whole project is based on..
  'BitBlt searches for a part of a picture, Cuts it and Pastes it to a specified location
  
'How to use BitBlt:
   'Example:  BitBlt TheDest.hDC, 0, 0, 71, 23, TheSource.hDC, 73, 200, SRCCOPY: TheDest.Refresh
   
   'TheDest.hdc -> is where the cut-out part of the picture will go
   '0,0 -> this tells BitBlt to start at left(x)-0 and top(y)-0
   '71,23 -> this tells BitBlt to take cut-out a width of: 71 and a height of:23 from the Source
   'TheSource(picturebox)-> contains the picture in whole
   '73,200 -> the x and y cordnates on the source picture. 71,23 above starts cutting at these cordnates
   'SRCCOPY -> Holds the cut-out part of the picture in memory to be stored in the "TheDest" picturebox
   'TheDest.Refresh -> Do this to reflect the SRCCOPY's cut-out into TheDest
   
   'That's it. It's pretty easy admit it. Just keep studying this until you get it
   'It will snap in your head trust me. It worked for me so it should for you.
   'NOTE: to get the x,y cordnates of a part in the picture as I have
   'USE: MSPAINT.EXE which comes with Win95-Win98 probobly NT but this is what
   'I used anyway. Goto: Star and Run and type in MSPAINT if you cant find it :)
   
   Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
  'This allows the AddIconToSysTray and RemoveIconFromSysTray functions to work
   Private Declare Function GetActiveWindow Lib "user32" () As Long
  'Gets the Active window open
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  'Tells a Window an objective and then manipulates it
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
  'Manipulates a object by sending a message to it's handle
    Private Declare Sub ReleaseCapture Lib "user32" ()
  'Calls the function ReleaseCapture in "user32" and Allows you to "drag" a form

'Consts:
    Private Const WM_RBUTTONUP = &H205
  'Sends the right mouse up call to an object or Handle
    Private Const WM_MOUSEMOVE = &H200
  'Detects mouse movements
    Private Const WM_LBUTTONDOWN = &H201
  'Sends the left mousebutton down call to an object or Handle
    Private Const WM_LBUTTONUP = &H202
  'Sends the left mousebutton up call to an object or Handle
    Private Const SRCCOPY = &HCC0020
  'Used in the BitBlt Function to store the Cut out picture in memory
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    Private Const HWND_TOPMOST = -1
    Private Const HWND_NOTOPMOST = -2
    Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  'Just some SetWindowPos Consts
    Private Const NIM_ADD = &H0
    Private Const NIM_DELETE = &H2
    Private Const NIF_ICON = &H2
    Private Const NIF_MESSAGE = &H1
    Private Const NIM_MODIFY = &H1
    Private Const NIF_TIP = &H4
    Private Const MAX_TOOLTIP As Integer = 64
    Private nfIconData As NOTIFYICONDATA
    Private Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uId              As Long
    uFlags           As Long
    ucallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
    End Type
   'The AddIconToSysTray/RemoveIconFromSysTray functions uses
   'These commands to add/remove the icon from the SystemTray

Private Function LoadSkin()
    BitBlt PicTitleBar.hDC, 0, 0, 360, 21, PicSource.hDC, 0, 1, SRCCOPY
'Cut's the TitleBar part out of the Pic and Copies it to PicTitleBar
    BitBlt PicBG.hDC, 0, 0, 350, 200, PicSource.hDC, 0, 23, SRCCOPY
'Cut's the BackGround part out of the Pic and Copies it to PicBG
    BitBlt PicMin.hDC, 0, 0, 17, 17, PicSource.hDC, 20, 230, SRCCOPY
'Cut's the MinimizeButtonUp part out of the Pic and Copies it to PicMin
    BitBlt PicExit.hDC, 0, 0, 17, 15, PicSource.hDC, 60, 230, SRCCOPY
'Cut's the ExitUp part out of the Pic and Copies it to PicExit
    BitBlt PicExitButton.hDC, 0, 0, 71, 23, PicSource.hDC, 217, 200, SRCCOPY
'Cut's the ExitButtonUp part out of the Pic and Copies it to PicExitButton
    BitBlt PicLoadButton.hDC, 0, 0, 71, 23, PicSource.hDC, 73, 200, SRCCOPY
'Cut's the ExitButtonUp part out of the Pic and Copies it to PicExitButton
    PicTitleBar.Refresh: PicBG.Refresh: PicMin.Refresh: PicExit.Refresh: PicExitButton.Refresh: PicLoadButton.Refresh
'Refreshes all the picture boxes to reflect off of the SRCCOPY
    Me.Height = 0
    Dim X As Integer
    For X = 1 To 2700
    Me.Height = Val(Me.Height) + 1
    Next X
'This is just to make a little effect on the skin when loaded
End Function

Private Sub Form_Load()
On Error GoTo here:
    PicSource = LoadPicture(App.Path & "\Skins\Skin1.bmp")
'Loads the skin pic into the PicSource PictureBox
    CenterForm Me
'Puts the form in the midle of the screen
    StayOnTop Me
'Tells the form to "be on top" of all of the other Windows
    LoadSkin
'Loads the default skin "\Skin1.bmp"
    SetObjectPositions
'Puts everything on the form to it's postion
Exit Sub
here:
    MsgBox "Please locate the folder " & """" & "Skins" & """" & " and place it in the same folder as the Executable.", vbSystemModal + vbCritical, "Error:": End
'If error, tell you then exit
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_LBUTTONDOWN
        RemoveIconFromSysTray
        Me.WindowState = 0
        Me.Show
        RemoveIconFromSysTray
    Case WM_RBUTTONUP
        Me.WindowState = 0
        Me.Show
        RemoveIconFromSysTray
    Case Else
End Select
'When the Form is in the systemtray. Detect that the left or right
'Mousebutton was clicked and removethe icon from systray and show the form
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then AddIconToSysTray Me, "Mike's Skin Example": Me.Hide
End Sub

Private Sub PicExit_Click()
    PicExitButton_Click
End Sub

Private Sub PicExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    BitBlt PicExit.hDC, 0, 0, 17, 15, PicSource.hDC, 40, 230, SRCCOPY: PicExit.Refresh
End Sub

Private Sub PicExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicExit.hDC, 0, 0, 17, 15, PicSource.hDC, 60, 230, SRCCOPY: PicExit.Refresh
End Sub

Private Sub PicExitButton_Click()
    Dim X As Long
    X& = MsgBox("Are you sure you want to exit?", vbSystemModal + vbInformation + vbYesNo, "Exit:"): If X& = vbYes Then End
End Sub

Private Sub PicExitButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicExitButton.hDC, 0, 0, 71, 23, PicSource.hDC, 145, 200, SRCCOPY: PicExitButton.Refresh
End Sub

Private Sub PicExitButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicExitButton.hDC, 0, 0, 71, 23, PicSource.hDC, 217, 200, SRCCOPY: PicExitButton.Refresh
End Sub

Private Sub PicLoadButton_Click()
On Error GoTo here:
    SendMessage SkinLoad.hWnd, WM_LBUTTONDOWN, 0, 0
        SendMessage SkinLoad.hWnd, WM_LBUTTONDOWN, 0, 0
            SendMessage SkinLoad.hWnd, WM_LBUTTONDOWN, 0, 0
            SendMessage SkinLoad.hWnd, WM_LBUTTONUP, 0, 0
        SendMessage SkinLoad.hWnd, WM_LBUTTONUP, 0, 0
    SendMessage SkinLoad.hWnd, WM_LBUTTONUP, 0, 0
  'This makes sure a Number in the LoadSkin combobox is selected
    PicSource = LoadPicture(App.Path & "\Skins\Skin" & SkinLoad.Text & ".bmp")
  'Loads Skin Picture into the PicSource picturebox to later use with BitBlt
    LoadSkin
  'Loads the skin
  Exit Sub
here: MsgBox "Please locate the folder " & """" & "Skins" & """" & " and place it in the same folder as the Executable.", vbSystemModal + vbCritical, "Error:": End
'if error tell you there is then exit
End Sub

Private Sub PicLoadButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    BitBlt PicLoadButton.hDC, 0, 0, 71, 23, PicSource.hDC, 1, 200, SRCCOPY: PicLoadButton.Refresh
End Sub

Private Sub PicLoadButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicLoadButton.hDC, 0, 0, 71, 23, PicSource.hDC, 73, 200, SRCCOPY: PicLoadButton.Refresh
End Sub

Private Sub PicMin_Click()
    Me.WindowState = 1
End Sub

Private Sub PicMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicMin.hDC, 0, 0, 17, 17, PicSource.hDC, 0, 230, SRCCOPY:    PicMin.Refresh
End Sub

Private Sub PicMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PicMin.hDC, 0, 0, 17, 17, PicSource.hDC, 20, 230, SRCCOPY:    PicMin.Refresh
End Sub
Private Sub PicTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 15: FormDrag Me: Me.MousePointer = 0
'This makes the form moved by calling the ReleaseCapture function
End Sub
Private Function StayOnTop(TheForm As Form)
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'Tells the form to be on top
End Function
Private Function NotOnTop(frm As Form)
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
'Tells the form to nt be on top
End Function
Private Function CenterForm(TheForm As Form)
    TheForm.Move (Screen.Width) / 2 - (TheForm.Width) / 2, (Screen.Height) / 2 - (TheForm.Height) / 2
'Moves the form to the center of the screen
End Function
Private Function FormDrag(TheForm As Form)
    ReleaseCapture
    SendMessage TheForm.hWnd, &HA1, 2, 0&
'Allows form to be able to be moved elsewhere instead of its titlebar
End Function
Private Function TimeOUT(HesitateTime)
Dim Hesitator As Long
    Hesitator = Timer
    Do While Timer - Hesitator < Val(HesitateTime)
    DoEvents
    Loop
'This pauses your form for a certain amount of time
End Function
Private Function AddIconToSysTray(TheForm As Form, MouseTipTitle As String)
    With nfIconData
        .hWnd = TheForm.hWnd: .uId = TheForm.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .ucallbackMessage = WM_MOUSEMOVE: .hIcon = TheForm.Icon.Handle
        .szTip = MouseTipTitle & vbNullChar: .cbSize = Len(nfIconData)
    End With
    Shell_NotifyIcon NIM_ADD, nfIconData
'Adds the form's icon into your System's Tray
End Function
Private Function RemoveIconFromSysTray()
    Shell_NotifyIcon NIM_DELETE, nfIconData
'Removes the form's icon from your SystemTray
End Function
Private Function SetObjectPositions()
    PicSource.Visible = False: PicTitleBar.Left = -10: PicMin.Left = 4150: PicExit.Left = 4400: EmailMSG.Left = 120
    DisplayMSG.Left = 120: SkinLoad.Left = 720: PicLoadButton.Left = 2330: PicExitButton.Left = 3480: PicBG.Left = -10
'Loads all the objects on the form to their Left position
    PicTitleBar.Top = -10: PicMin.Top = 20: PicExit.Top = 20: DisplayMSG.Top = 120: SkinLoad.Top = 1965
    PicLoadButton.Top = 2200: PicExitButton.Top = 2200: PicBG.Top = 300: EmailMSG.Top = 1000
'Loads all the objects on the form to their Top position
End Function

'<END> of coding
