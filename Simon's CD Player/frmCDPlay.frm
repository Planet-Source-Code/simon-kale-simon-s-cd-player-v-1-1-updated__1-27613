VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCdplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simon's CD Player"
   ClientHeight    =   2415
   ClientLeft      =   4590
   ClientTop       =   2685
   ClientWidth     =   4515
   ForeColor       =   &H00008080&
   Icon            =   "frmCDPlay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4515
   Begin VB.Timer CheckDeviceReady 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   1320
   End
   Begin VB.Timer CDTimer 
      Interval        =   1000
      Left            =   3360
      Top             =   1320
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   525
      Left            =   430
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "[00] 00:00"
      Top             =   270
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3720
   End
   Begin VB.ComboBox cboTrack 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCDPlay.frx":030A
      Left            =   720
      List            =   "frmCDPlay.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ComboBox cboArtist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdEject 
      Height          =   375
      Left            =   4080
      Picture         =   "frmCDPlay.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eject"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdNextTrack 
      Height          =   375
      Left            =   3720
      Picture         =   "frmCDPlay.frx":0458
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next Track"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdFastForward 
      Height          =   375
      Left            =   3360
      Picture         =   "frmCDPlay.frx":05A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Skip Forwards"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdRewindBack 
      Height          =   375
      Left            =   3000
      Picture         =   "frmCDPlay.frx":06EC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Skip Backwards"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdPreviousTrack 
      Height          =   375
      Left            =   2640
      Picture         =   "frmCDPlay.frx":0836
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Previous Track"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Height          =   375
      Left            =   4080
      Picture         =   "frmCDPlay.frx":0980
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Stop"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Height          =   375
      Left            =   3720
      Picture         =   "frmCDPlay.frx":0ACA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pause"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      Picture         =   "frmCDPlay.frx":0C14
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Play"
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar stbCDPlayer 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   2100
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3942
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3942
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTrack 
      Caption         =   "Trac&k :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblTitleMsg 
      Caption         =   "Title :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1340
      Width           =   615
   End
   Begin VB.Label lblArtist 
      Caption         =   "&Artist :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Menu mnuDisc 
      Caption         =   "&Disc"
      Begin VB.Menu mnuDiscEditPlayList 
         Caption         =   "Edit Play &List..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiscExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDiscTrackInfo 
         Caption         =   "&Disc/Track Info"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTrackTimeElapsed 
         Caption         =   "Track Time &Elapsed"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewTrackTimeRemaining 
         Caption         =   "Track Time &Remaining"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDiscTimeRemaining 
         Caption         =   "Dis&c Time Remaining"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewVolumeControl 
         Caption         =   "&Volume Control"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opitons"
      Begin VB.Menu mnuOptionsRandomOrder 
         Caption         =   "&Random Order"
      End
      Begin VB.Menu mnuOptionsContinuousPlay 
         Caption         =   "&Continuous Play"
      End
      Begin VB.Menu mnuOptionsIntroPlay 
         Caption         =   "&Intro Play"
      End
      Begin VB.Menu mnuOptionsbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsPreferences 
         Caption         =   "&Preferences"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelpTopics 
         Caption         =   "&Help Topics"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSimonsCDPlayer 
         Caption         =   "&About Simon's CD Player"
      End
   End
End
Attribute VB_Name = "frmCdplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'**********************************************************************************
'Simon's CD Player
'Written By: Simon Kale
'EMail: simonkale@yahoo.com
'Date: 28th September 2001
'Website: http://vbunlimited.cjb.net

'This is a MS Windows CD Player Clone. :) (not complete though)
'Copyright(c)2001 - 2002. All Rights Reserved.

'Your suggestions and comments are welcome.
'Please send bug reports to: simonkale@yahoo.com

'Happy learning and coding...........

'**********************************************************************************
'**********************************************************************************

Option Explicit
'now "snd" can be used to call the sub routine that resides in "clsCDPlayer"
Dim snd As clsCDPlayer
'returns True if CD door is open otherwise False
Dim blnDoorOpen As Boolean
'returns True if combo box's ("cboTrack") click event if to be triggered else False
Dim blnSkipcboTrack As Boolean
'gets the track number
Dim intTimeTrack As Integer
'gets the track minute
Dim intTimeMinute As Integer
'gets the track second
Dim intTimeSecond As Integer
'Address of a buffer that receives return information from MCI Command String
Dim strReturnString As String * 40
'gets the return values of "mciSendString"
Dim lngReturnValues As Long

Private Sub cboTrack_Click()
    '* Purpose: Changes the CD track to the track that is selected by user
    
    'check if blnSkipcboTrack = True
    'if so exit the subroutine
    If blnSkipcboTrack = True Then
        GoTo Proc_End
    Else
        'change the track number to the selected one
        'check if an audio cd is loaded
        If snd.CheckCD = True Then
            'check if it is not a Data CD
            If snd.GetTotalNumberOfTracks >= 2 Then
                'check if CD is playing
                If snd.CheckIfCDIsPlaying = True Then
                    'get the CD position
                    snd.GetCDPosition
                    'seek to the track selected by the user
                    lngReturnValues = mciSendString("seek cd to " & _
                                    CInt(Mid$(cboTrack.Text, 7, 2)), 0, 0, 0)
                    'if the previous track was playing then play the currently
                    'selected track
                    lngReturnValues = mciSendString("play cd from " & _
                                CInt(Mid$(cboTrack.Text, 7, 2)), 0, 0, 0)
                    'disable play and disable stop and pause buttons
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                Else
                    'otherwise just seek to selected track but don't play it
                    lngReturnValues = mciSendString("seek cd to " & _
                                    CInt(Mid$(cboTrack.Text, 7, 2)), 0, 0, 0)
                End If
            End If
        End If
    End If

Proc_End:
    Exit Sub
    
End Sub

Private Sub CDTimer_Timer()
    '* Purpose: Updates the track information every seconds
    
    'check if cd is loaded
    If snd.CheckCD = True Then
        'update the information every second
        UpdateValues
    Else
        'disable the timer
        CDTimer.Enabled = False
        'update values
        Call CheckAudioCD
        'enable another timer which check if the CD is loaded now
        CheckDeviceReady.Enabled = True
        CheckDeviceReady.Interval = 1000
        'Call the routine which check if the CD is loaded
        Call CheckIfDeviceIsReady
    End If
    
End Sub

Private Sub CheckDeviceReady_Timer()
    '* Purpose: Checks if CD is loaded now or if Device is ready
    
    Call CheckIfDeviceIsReady
End Sub

Private Sub cmdEject_Click()
    '* Purpose: If CD door is close then open it and vice-versa
    
    'eject CD drawer
    'check if cd door is opened
    If blnDoorOpen = True Then
        'close the cd door
        snd.EjectCloseCD
        'set blnDoorOpen to False
        blnDoorOpen = False
    Else
        'open the cd door
        snd.EjectOpenCD
        'set the blnDoorOpen to True
        blnDoorOpen = True
        'disable the timer
        CDTimer.Enabled = False
        'update values
        Call CheckAudioCD
        'enable another timer which check if the CD is loaded now
        CheckDeviceReady.Enabled = True
        CheckDeviceReady.Interval = 1000
        'Call the routine which check if the CD is loaded
        Call CheckIfDeviceIsReady
    End If
End Sub

Private Sub cmdFastForward_Click()
    '* Purpose: Fast Forwards the current track by 10 seconds
    
    'fast forward the current track
    Dim strReturnString As String * 30
    Dim lngForwardValue As Long
    'set the time format to Milliseconds
    snd.SetCDFormat_MilliSeconds
    'fast forward the current track
    mciSendString "status cd position wait", strReturnString, Len(strReturnString), 0
    'forward the track position by 10 secs
    '10 Seconds = 10 * 1000
    lngForwardValue = CStr(CLng(strReturnString)) + 10 * 1000
    'check if cd is playing
    If snd.CheckIfCDIsPlaying = True Then
        mciSendString "seek cd to " & lngForwardValue, 0, 0, 0
        'play the cd from new position
        mciSendString "play cd from " & lngForwardValue, 0, 0, 0
    Else
        'if not playing then just change the cd position by 10 secs but don't play
        mciSendString "seek cd to " & lngForwardValue, 0, 0, 0
    End If
    'set the time format to TMSF
    snd.SetCDFormat_TMSF
    'update the track combobox "cboTrack"
    UpdateCboTrackValue
End Sub

Private Sub cmdNextTrack_Click()
    '* Purpose: Play the next track on CD
    
    'select the next track
    Dim CurrentTrack As Integer
    'get the current track
    CurrentTrack = snd.GetCurrentTrack
        'check if current track is the last track
        If CurrentTrack = snd.GetTotalNumberOfTracks Then
            'set the current track to zero i.e initial
            CurrentTrack = 0
        End If
            'check if CD is playing
            If snd.CheckIfCDIsPlaying = True Then
                'change the cd position to the first track
                'new cd position = CurrentTrack + 1 (i.e. 0 + 1 = 1)
                mciSendString "seek cd to " & CurrentTrack + 1, 0, 0, 0
                'play the track
                mciSendString "play cd from " & CurrentTrack + 1, 0, 0, 0
            Else
                'seek to next track but don't play it
                mciSendString "seek cd to " & CurrentTrack + 1, 0, 0, 0
                'disable stop and pause and enable play buttons
                cmdStop.Enabled = False
                cmdPause.Enabled = False
                cmdPlay.Enabled = True
                'set focus to play button
                cmdPlay.SetFocus
            End If
            'update the track combobox "cboTrack"
            UpdateCboTrackValue
End Sub

Private Sub cmdPause_Click()
    '* Purpose: Pauses the currently playing track
    
    'pause the current track
    snd.PauseTrack
    'disable pause and enable play buttons
    cmdPause.Enabled = False
    cmdPlay.Enabled = True
    'set focus to play button
    cmdPlay.SetFocus
End Sub

Private Sub cmdPlay_Click()
    '* Purpose: Plays the currently selected track
    
    'start playing the selected track
    snd.PlayTrack
    'disable play and enable stop and pause buttons
    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    cmdPause.Enabled = True
    'set focus to stop button
    cmdPause.SetFocus
End Sub

Private Sub cmdPreviousTrack_Click()
    '* Purpose: Changes the current track to the track before it
    
    'select the previous track
    Dim CurrentTrack As Integer
    'get the current track
    CurrentTrack = snd.GetCurrentTrack
        'check if it is the first track
        If CurrentTrack = 1 Then
            'if first track then the previous track is the last track on the cd
            CurrentTrack = snd.GetTotalNumberOfTracks
                'for some unknown reason, we have to add 1 to
                'the total no. of tracks and then subtrack 1 again to make
                'the selection of the last track
                CurrentTrack = CurrentTrack + 1
        End If
        'check if CD is playing
        If snd.CheckIfCDIsPlaying = True Then
            mciSendString "seek cd to " & CurrentTrack - 1, 0, 0, 0
            'play the track
            mciSendString "play cd from " & CurrentTrack - 1, 0, 0, 0
        Else
            'seek to next track but don't play it
            mciSendString "seek cd to " & CurrentTrack - 1, 0, 0, 0
            'disable stop and pause and enable play buttons
            cmdStop.Enabled = False
            cmdPause.Enabled = False
            cmdPlay.Enabled = True
            'set focus to play button
            cmdPlay.SetFocus
        End If
        'update the track combobox "cboTrack"
        UpdateCboTrackValue
End Sub

Private Sub cmdRewindBack_Click()
    '* Purpose: Rewinds the current track by 10 seconds
    
    'rewind the current track
    Dim RewindSpeed As Long
    Dim RewindValue As Long
    RewindSpeed = 10
    'set the time format to Milliseconds
    snd.SetCDFormat_MilliSeconds
    'rewind the current track
    mciSendString "status cd position wait", strReturnString, Len(strReturnString), 0
    'rewind the track position by 10 secs
    RewindValue = CStr(CLng(strReturnString)) - RewindSpeed * 1000
    If snd.CheckIfCDIsPlaying = True Then
        mciSendString "seek cd to " & RewindValue, 0, 0, 0
        'play the cd from new position
        mciSendString "play cd from " & RewindValue, 0, 0, 0
    Else
        'if not playing then just change the cd position by 10 secs but don't play
        mciSendString "seek cd to " & RewindValue, 0, 0, 0
    End If
    'set the time format to TMSF
    snd.SetCDFormat_TMSF
    'update the track combobox "cboTrack"
    UpdateCboTrackValue
End Sub

Private Sub cmdStop_Click()
    '* Purpose: Stops the currently playing track
    
    'stop playing the track
    snd.StopTrack
    snd.ReadyDevice
    'disable stop and pause and enable play buttons
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    cmdPlay.Enabled = True
    'set focus to play button
    cmdPlay.SetFocus
    'update the track combobox "cboTrack"
    UpdateCboTrackValue
End Sub

Private Sub Form_Load()
    'check if previous instance of the program is opened
    If App.PrevInstance = True Then
        'if so then exit
        End
        Exit Sub
    End If
    'initialise "clsCDPlayer" class module as "snd"
    Set snd = New clsCDPlayer
        'check if cdplayer device is ready
        If snd.ReadyDevice = False Then
            'call the device not read subroutine
            Call DeviceNotReady
        Else
            'check if cd is present in the drive
            CheckAudioCD
                'check if CD is present in the drive
                If snd.CheckCD = True Then
                    'update the track combobox "cboTrack"
                    UpdateCboTrackValue
                Else
                    'enable another timer which check if the CD is loaded now
                    CheckDeviceReady.Enabled = True
                    CheckDeviceReady.Interval = 1000
                End If
        End If
        'this is to insure we get different track every time
        Randomize
        'check if the program's registry key alreadly exits
        If OpenRegistrySettings() = False Then
            'if not, create the registry settings for the program
            Call CreateRegistrySettings
        End If
        'load the FormLoad event of "Preferences" dialog box
        Call frmPreferences.Form_Load
        'save the menuitems states in the Registry
        Call GetMenuStateFromRegistry
        'get main window Top and Left position
        Call RegGetFormDimension
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '* Purpose: Calls the mnuDiscExit menu item
    
    'exit the program, unloading all the opened devices
    mnuDiscExit_Click
End Sub

Private Sub mnuDiscExit_Click()
    '* Purpose: Exits the program unloading all the opened devices
    
    'check if "Stop CD Playing On Exit" is checked in "Preferences" dialog box
    'is checked
    If frmPreferences.chkStopCDPlay = Checked Then
        'stop playing CD before exiting
        snd.StopTrack
    End If
    'save the menuitems states in the Registry
    'first open "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    Call OpenRegistrySettings
        If OpenRegistrySettings = True Then
            Call RegRandomOrder
            Call RegContinuousPlay
            Call RegIntroPlay
            Call CloseRegistry
        End If
    'save main windows position
    Call RegFormLeftAndTop
    'close all the devices
    snd.UnloadAll
    'close all windows of the program
    Unload frmAbout
    Unload frmPreferences
    Unload Me
End Sub

Private Sub mnuHelpSimonsCDPlayer_Click()
    '* Purpose: Shows the about dialog box
    
    'show about dialog box
    frmAbout.Show
End Sub

Private Sub mnuOptionsContinuousPlay_Click()
    '* Purpose: Checks or unchecks the "mnuOptionsContinuousPlay" menuitem
    
    'check or uncheck the continuous play menu item
    mnuOptionsContinuousPlay.Checked = Not mnuOptionsContinuousPlay.Checked
End Sub

Private Sub mnuOptionsIntroPlay_Click()
    '* Purpose: Plays tracks intro only for specified seconds in preference dialod box
    
    'check or uncheck the "Intro Play" menuitem
    mnuOptionsIntroPlay.Checked = Not mnuOptionsIntroPlay.Checked
    
End Sub

Private Sub mnuOptionsPreferences_Click()
    '* Purpose: Shows the Preferences dialog box
    
    'show preferences dialog box
    frmPreferences.Show
End Sub

Private Sub mnuOptionsRandomOrder_Click()
    '* Purpose: Randomply plays the track if corresponding menuitem is checked
    
    'check or uncheck the random order menuitem
    mnuOptionsRandomOrder.Checked = Not mnuOptionsRandomOrder.Checked
End Sub

Private Sub mnuViewDiscTrackInfo_Click()
    '* Purpose: Shows or Hides "cboArtist", "txtTitle" and "cboTrack" windows

    'check or uncheck the disc/track info menu
    mnuViewDiscTrackInfo.Checked = Not mnuViewDiscTrackInfo.Checked
        'check if "mnuViewDiscTrackInfo" is checked
        If mnuViewDiscTrackInfo.Checked = False Then
            'check if "mnuViewStatusbar" is checked
            If mnuViewStatusbar.Checked = False Then
                'decrease the main window height
                'main window default height = 3075
                'cboArtist + txtTitle + cboTrack height = 990
                'status bar height = 315
                frmCdplay.Height = 3075 - 990 - 315
            Else
                frmCdplay.Height = 3075 - 990
            End If
            'hide cboArtist, txtTitle and cboTrack windows
            cboArtist.Visible = False
            txtTitle.Visible = False
            cboTrack.Visible = False
            'hide disc/track info
            lblArtist.Visible = False
            lblTitleMsg.Visible = False
            lblTrack.Visible = False
        Else
            If mnuViewStatusbar.Checked = False Then
                frmCdplay.Height = 3075 - 315
            Else
                frmCdplay.Height = 3075
            End If
            'show cboArtist, txtTitle and cboTrack windows
            cboArtist.Visible = True
            txtTitle.Visible = True
            cboTrack.Visible = True
            'hide disc/track info
            lblArtist.Visible = True
            lblTitleMsg.Visible = True
            lblTrack.Visible = True
        End If
End Sub

Private Sub mnuViewStatusbar_Click()
    '* Purpose: Shows or Hides status bar
    
    'check or uncheck the statusbar menu
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
        'decrease/increase the main window height
        If mnuViewStatusbar.Checked = False Then
            If mnuViewDiscTrackInfo.Checked = False Then
                'decrease the main window height
                'main window default height = 3075
                'cboArtist + txtTitle + cboTrack height = 990
                'status bar height = 315
                frmCdplay.Height = 3075 - 990 - 315
            Else
                frmCdplay.Height = 3075 - 315
            End If
            'show the status bar
            stbCDPlayer.Visible = False
        Else
            If mnuViewDiscTrackInfo.Checked = False Then
                frmCdplay.Height = 3075 - 990
            Else
                frmCdplay.Height = 3075
            End If
            'show the status bar
            stbCDPlayer.Visible = True
        End If
End Sub

Private Function CheckAudioCD() As Boolean
    '* Purpose: 1) Checks if CD is loaded
    '*          2) If loaded then is it an audio CD
    '*          3) If an audio CD then updates combo boxes with corresponding values
    '*          4) Enables Timer and set the interval
    '*          5) Enables buttons
    '*          6) Shows the total CD length in status bar
    
    'this routine check if CD is present in the CD-ROM Drive
        If snd.CheckCD = True Then
            cboArtist.Clear
            'the following code uses "FindCDDriveLetter" function to
            'extract CD Drive letter
            cboArtist.AddItem "New Artist" & Space$(35) & "<" & FindCDDriveLetter & ">"
            cboArtist.Text = cboArtist.List(0)
            txtTitle.Text = "New Title"
            'update combo box "cboTrack" with track numbers
            UpdateComboBoxTrack
            'enable timer and set the interval to 1 sec
            CDTimer.Enabled = True
            CDTimer.Interval = 1000
            'call the timer
            CDTimer_Timer
                'check if CD is playing
                'if playing then disable play button
                If snd.CheckIfCDIsPlaying = True Then
                    'disable play and enable pause and stop buttons
                    cmdPlay.Enabled = False
                    cmdPause.Enabled = True
                    cmdStop.Enabled = True
                Else
                    'disable pause and stop buttons and enable rest
                    cmdPlay.Enabled = True
                    cmdPause.Enabled = False
                    cmdStop.Enabled = False
                    cmdPreviousTrack.Enabled = True
                    cmdRewindBack.Enabled = True
                    cmdFastForward.Enabled = True
                    cmdNextTrack.Enabled = True
                End If
                'set total cd length on the statusbar panel 1
                Dim TotalCDTime As String
                Dim Minute As String
                'get the cd total length
                TotalCDTime = snd.GetCDLength
                'only display total hours and minute on the status bar
                stbCDPlayer.Panels(1) = "Total Play: " & Mid$(TotalCDTime, 1, 5) _
                                    & " m:s"
        Else
            'disable all buttons except eject button
            cmdPlay.Enabled = False
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            cmdPreviousTrack.Enabled = False
            cmdRewindBack.Enabled = False
            cmdFastForward.Enabled = False
            cmdNextTrack.Enabled = False
            cboArtist.Clear
            cboArtist.AddItem "Data or no disc loaded"
            cboArtist.Text = cboArtist.List(0)
            txtTitle.Text = "Please insert an audio compact disc"
            cboTrack.Clear
            'set total cd length on the statusbar panel 1
            stbCDPlayer.Panels(1) = "Total Play: 00:00" & " m:s"
            'set current track length on the statusbar panel 2
            stbCDPlayer.Panels(2) = "Track: 00:00" & " m:s"
            'update the track time and show on txtTime
            txtTime.Text = "[00] 00:00"
            'disable timer and set the interval to 0
            CDTimer.Enabled = False
            CDTimer.Interval = 0
        End If
End Function

Private Sub UpdateComboBoxTrack()
    '* Purpose: Loads the CD track number to the combo box "cboTrack"
    
    'fill the combobox "cboTrack" with track numbers
    Dim TotalNumOfTracks As Integer
    'first check if it is a data CD
    'if only one track then data cd
    'this won't play a cd containing only one track (single CDs)
    'get the total number of tracks
    TotalNumOfTracks = snd.GetTotalNumberOfTracks
    'check if it is less than equal to 1
    If TotalNumOfTracks <= 1 Then
        'if it is then
        'disable all buttons except eject
        cmdPlay.Enabled = False
        cmdPause.Enabled = False
        cmdStop.Enabled = False
        cmdPreviousTrack.Enabled = False
        cmdRewindBack.Enabled = False
        cmdFastForward.Enabled = False
        cmdNextTrack.Enabled = False
        cboArtist.AddItem "Disc or Data no loaded"
        cboArtist.Text = cboArtist.List(0)
    Else
        'fill the "cboTrack" with track numbers
        Dim I As Integer
        'clear the content of the combo box
        cboTrack.Clear
            For I = 1 To TotalNumOfTracks
                Dim intSR As Integer
                'if track number is greater than or equal to 10 then
                'space required to right align the track number is "37"
                If I >= 10 Then
                    intSR = 37
                    'add the track number to combobox
                    cboTrack.AddItem "Track " & I & Space$(intSR) & "<" & I & ">"
                Else
                    'otherwise, it is "40"
                    intSR = 39
                    'add the track number to combobox
                    cboTrack.AddItem "Track " & I & Space$(intSR) & "<0" & I & ">"
                End If
                'if I is equal to total number of tracks on the cd then exit loop
                If I = TotalNumOfTracks Then Exit For
            Next I
            'skip the combo box "cboTrack" click event
            blnSkipcboTrack = True
            'set the combobox text to display first item i.e. track 1
            cboTrack.Text = cboTrack.List(0)
            'don't skip the combo box "cboTrack" click event
            blnSkipcboTrack = False
    End If
End Sub

Private Sub UpdateValues()
    '* Purpose: Updates combo box and various windows information from CD
    
    'update the values
    '*****************
    'update txtTime
    'display the track length on "txtTime"
    'get the track number
    intTimeTrack = CInt(Mid$(snd.GetCDPosition, 1, 2))
    'get the track minute
    intTimeMinute = CInt(Mid$(snd.GetCDPosition, 4, 2))
    'get the track second
    intTimeSecond = CInt(Mid$(snd.GetCDPosition, 7, 2))
    'display it in the time text box
    txtTime.Text = "[" & Format(intTimeTrack, "00") & "] " _
                    & Format(intTimeMinute, "00") & ":" & Format(intTimeSecond, "00")
        'check if track minute = "0" and second is equal to "1"
        'if they are equal then probably the current track has finished
        'so update the combo box to update the track number
        If intTimeMinute = "0" And intTimeSecond = "1" Then
            'update the track combobox "cboTrack"
            UpdateCboTrackValue
        End If
        'continously play the track
        Call ContinuousPlay
        'randomly play the tracks
        Call RandomlyPlayTrack
        'play track intro only
        Call PlayTrackIntro
        'set current track length on the statusbar panel 2
        Dim TotalTrackLength As String
        'get the current track length
        TotalTrackLength = snd.GetCurrentTrackLength
        'display only minute ans seconds on the status bar
        stbCDPlayer.Panels(2) = "Track: " & Mid$(TotalTrackLength, 1, 5) & " m:s"
End Sub

Private Sub UpdateCboTrackValue()
    '* Purpose: Selects the "Current Track - 1" as the current track in combo box
    
    'update the track number "cboTrack"
    'skip the combo box "cboTrack" click event
    blnSkipcboTrack = True
    'the combo box index start from 0
    'so we have to subtract 1 from the current track number
    cboTrack.Text = cboTrack.List(snd.GetCurrentTrack - 1)
    'don't skip the combo box "cboTrack" click event
    blnSkipcboTrack = False
End Sub

Private Sub mnuViewVolumeControl_Click()
    '* Purpose: Shows MS Windows Volume control
    
    'show the windows volume control
    Dim strPath As String, strSave As String, strProgramName As String
    Dim lngOpenVolumeControl As Long
    'Create a buffer string
    strSave = String(200, Chr$(0))
    'Get the windows directory and append '\SndVol32.exe' to it
    strPath = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave)))
    'combine path and program name
    strProgramName = strPath & "\sndvol32.exe"
    'open the "sndvol32.exe" program (volume control)
    lngOpenVolumeControl = ShellExecute(Me.hwnd, "open", strProgramName, 0, 0, _
                        SW_SHOWNORMAL)
End Sub

Private Sub DeviceNotReady()
    '* Purpose: Since the device is not ready it display the "CDTimer" timer so that
    '*          it can update any CD information and enables "CheckDeviceReady"
    '*          timer to check if the device is ready.
    
    'disable all the buttons
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdPreviousTrack.Enabled = False
    cmdRewindBack.Enabled = False
    cmdFastForward.Enabled = False
    cmdNextTrack.Enabled = False
    cmdEject.Enabled = False
    
    'clear the combo box
    cboArtist.Clear
    'add string to combo box
    cboArtist.AddItem "This drive is in use"
    cboArtist.Text = cboArtist.List(0)
    'set the string to text box
    txtTitle.Text = "Waiting for the drive to become available"

    'set total cd length on the statusbar panel 1
    stbCDPlayer.Panels(1) = "Total Play: 00:00" & " m:s"
    'set current track length on the statusbar panel 2
    stbCDPlayer.Panels(2) = "Track: 00:00" & " m:s"
    'update the track time and show on txtTime
    txtTime.Text = "[00] 00:00"
    
    'disable "CDTimer" timer
    CDTimer.Enabled = False
    'enable "CheckDeviceReady" timer
    CheckDeviceReady.Enabled = True
End Sub

Private Sub ContinuousPlay()
    '* Purpose: Continuously play the track if "mnuOptionsContinuousPlay" is checked
        
        'check if current track is the last track on cd
        If intTimeTrack = snd.GetTotalNumberOfTracks Then
            'if last track then get the total length of the last track
            mciSendString "status cd length track " _
                        & snd.GetTotalNumberOfTracks, strReturnString, _
                        Len(strReturnString), 0
            'check if trackminute is equal to the last track minute
            'and tracksecond is equal to "0"
            If intTimeMinute = CInt(Mid$(strReturnString, 1, 2)) And _
                        intTimeSecond = "0" Then
                'check if continuous play is checked
                If mnuOptionsContinuousPlay.Checked = True Then
                    'if checked then play first track
                    mciSendString "play cd from 1", 0, 0, 0
                Else
                    'else seek to first track but don't play the track
                    mciSendString "seek cd to 1", 0, 0, 0
                    'disable stop and pause and enable play buttons
                    cmdStop.Enabled = False
                    cmdPause.Enabled = False
                    cmdPlay.Enabled = True
                    'set focus to play button
                    cmdPlay.SetFocus
                    'update the track combobox "cboTrack"
                    UpdateCboTrackValue
                End If
            End If
        End If
End Sub

Private Sub RandomlyPlayTrack()
    '* Purpose: Randomly plays the track if menuitem "Random Order" is checked
    
    'check if "Random Order" is check or not
    'if checked then randomly play the tracks
    If mnuOptionsRandomOrder.Checked = True Then
            'check if we have reached to the end of the track
            If intTimeMinute = Int(Mid$(snd.GetCurrentTrackLength, 2, 1)) And _
                intTimeSecond = Int(Mid$(snd.GetCurrentTrackLength, 4, 2)) Then
                    'then randomly selected the next track
                    Dim intRandomTrack As Integer
                    'the formula to generate random number is:
                    'I = Int(lmax - lmin + 1) * Rnd) + lmin
                    intRandomTrack = Int((snd.GetTotalNumberOfTracks - 1) * Rnd) + 1
                        'check if CD is playing
                        If snd.CheckIfCDIsPlaying = True Then
                            'seek track to next random track and play it
                            lngReturnValues = mciSendString("seek cd to " _
                                            & intRandomTrack, 0, 0, 0)
                            'play the track
                            lngReturnValues = mciSendString("play cd from " _
                                            & intRandomTrack, 0, 0, 0)
                        Else
                            'seek track to next random track but don't play it
                            lngReturnValues = mciSendString("seek cd to " _
                                            & intRandomTrack, 0, 0, 0)
                        End If
            End If
    End If
End Sub

Private Sub PlayTrackIntro()
    '* Purpose: Plays tracks intro only for specified seconds in preference dialog box
    
    'check if "Intro Play" is checked
    If mnuOptionsIntroPlay.Checked = True Then
        'play tracks intro only
        Dim intCDTrack As Integer
        intCDTrack = snd.GetCurrentTrack
            'get the number of seconds to play from preferences dialog box
            If intTimeSecond >= CInt(frmPreferences.txtIntroPlay.Text) Then
                'check if it is the last track on CD
                If intCDTrack = snd.GetTotalNumberOfTracks Then
                    'seek track to first track
                    lngReturnValues = mciSendString("seek cd to 1", 0, 0, 0)
                Else
                    'seek track to next random track and play it
                    lngReturnValues = mciSendString("seek cd to " _
                                    & intCDTrack + 1, 0, 0, 0)
                End If
                'play the track
                lngReturnValues = mciSendString("play cd from " _
                                & intCDTrack + 1, 0, 0, 0)
            End If
    End If
End Sub

Private Sub CheckIfDeviceIsReady()
    '* Purpose: check it CD is loaded or Device is ready
    
    'check if CD is loaded
    If snd.CheckCD = True Then
        'check if it is an Audio CD
        If snd.GetTotalNumberOfTracks >= 2 Then
            'enable time and set its interval to 1 second
            CDTimer.Enabled = True
            CDTimer.Interval = 1000
            'disable "ChechDeviceReady"
            CheckDeviceReady.Enabled = False
            'Updates the values
            Call CheckAudioCD
        End If
    End If
End Sub
