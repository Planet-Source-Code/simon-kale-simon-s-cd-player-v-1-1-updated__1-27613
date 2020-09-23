VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'A buffer that receives return information from "mciSendString"
Dim strReturnString As String * 40
'gets the return values of "mciSendString"
Dim lngReturnValues As Long

Public Function PlayTrack()
    '* Purpose: Plays the selected track
        
    'Play the currrent/selected track
    lngReturnValues = mciSendString("play cd", 0, 0, 0)
End Function

Public Function StopTrack()
    '* Purpose: Stops the currently playing track
    
    'Stop the currently playing track
    lngReturnValues = mciSendString("stop cd", 0, 0, 0)
End Function

Public Function EjectOpenCD()
    '* Purpose: opens the CD door
    
    'Eject the CD door
    lngReturnValues = mciSendString("set cd door open", 0, 0, 0)
End Function

Public Function ReadyDevice() As Boolean
    '* Purpose: Intializes the CD
        
    'unloads the CD and all associated resources
    UnloadAll
        'Ready the device
        'if not ready then set ReadyDevice to False otherwise True
        If SetCDPlayerReady = False Then
            'set False
            ReadyDevice = False
        Else
            'otherwise True
            ReadyDevice = True
        End If
End Function

Public Function SetCDPlayerReady() As Boolean
    '* Purpose: Opens the CD device for playing
    
    'Set the CD Player ready
    lngReturnValues = mciSendString("open cdaudio alias cd wait shareable", 0, 0, 0)
        'check if error occured
        If lngReturnValues = 0 Then
            'then set True
            SetCDPlayerReady = True
        Else
            'otherwise set False
            SetCDPlayerReady = False
        End If
End Function

Public Function UnloadAll()
    '* Purpose: Closes all the devices opened by this program
    
    'close all the devices
    lngReturnValues = mciSendString("close all", 0, 0, 0)
End Function

Public Function PauseTrack()
    '* Purpose: Pauses the currrently playing track
    
    'Pause the currrently playing track
    lngReturnValues = mciSendString("pause cd", 0, 0, 0)
End Function

Public Function EjectCloseCD()
    '* Purpose: Closes the CD door
    
    'Close the CD door
    lngReturnValues = mciSendString("set cd door closed", 0, 0, 0)
End Function

Public Function CheckCD() As Boolean
    '* Purpose: Checks if CD is present in the CD-ROM drive
    '* Returns: CheckCD = True if CD is present in the drive otherwise False
    
    'check if CD is present in the drive
    lngReturnValues = mciSendString("status cd media present", strReturnString, _
                    Len(strReturnString), 0)
        'if no error occured then
        If lngReturnValues = 0 Then
            'check if CD is present
            If (CBool(strReturnString)) = True Then
                'set "CheckCD" to True
                CheckCD = True
            Else
                'if CD is not present
                'then set "CheckCD" to False
                CheckCD = False
            End If
        End If
End Function

Public Function GetCDLength() As String
    '* Purpose: Gets the total length of the CD loaded
    '* Returns: The total length of the CD in Track:Minutes:Seconds:Frames format
    
    'this function returns the total length of the CD
    lngReturnValues = mciSendString("status cd length wait", strReturnString, _
                    Len(strReturnString), 0)
        'if no error occured then
        If lngReturnValues = 0 Then
            'save the length in "GetCDLength"
            GetCDLength = strReturnString
        End If
End Function

Public Function CheckIfCDIsPlaying() As Boolean
    '* Purpose: Check if CD is playing or not
    '* Returns: True if playing else False
    
    'check if CD is playing
    lngReturnValues = mciSendString("status cd mode", strReturnString, _
                    Len(strReturnString), 0)
        'if no error occured then
            If lngReturnValues = 0 Then
                'check if CD is playing
                'if playing then set "CheckIfCDIsPlaying" to True
                If Mid$(strReturnString, 1, 7) = "playing" Then
                    CheckIfCDIsPlaying = True
                End If
            End If
End Function

Public Function SetCDFormat_MilliSeconds()
    '* Purpose: Sets the CD time format to MilliSeconds
    
    'set the time format to milliseconds abbreviated as "m:s"
    lngReturnValues = mciSendString("set cd time format milliseconds", 0, 0, 0)
End Function

Public Function SetCDFormat_TMSF()
    '* Purpose: Sets the CD time format to Track:Minutes:Seconds:Frames
    
    'set the time format to track, minutes, seconds, and frames
    lngReturnValues = mciSendString("set cd time format tmsf", 0, 0, 0)
End Function

Public Function GetMCIErrorString(lngReturnValues As Long) As String
    '* Purpose: Retrieves a string that describes the specified MCI error code
    '* Accepts: lngReturnValues - Error code returned by the mciSendString function
    '* Returns: The MCI error message
    
    'create a buffer
    GetMCIErrorString = Space$(255)
    'retrieve the error string
    mciGetErrorString lngReturnValues, GetMCIErrorString, Len(GetMCIErrorString)
    'strip off the trailing spaces
    GetMCIErrorString = Trim$(GetMCIErrorString)
End Function

Public Function GetCurrentTrackLength() As String
    '* Purpose: Gets the total length of the current track
    '* Returns: the current track length
    
    Dim CurrentTrack As Integer
    'get the current track
    CurrentTrack = GetCurrentTrack()
    'get the length of the current track
    lngReturnValues = mciSendString("status cd length track " & CurrentTrack, _
                    strReturnString, Len(strReturnString), 0)
        'if no error occured then
        If lngReturnValues = 0 Then
            'save the current track length in "GetCurrentTrackLength"
            GetCurrentTrackLength = strReturnString
        End If
End Function

Public Function GetCurrentTrack() As Integer
    '* Purpose: Gets the currently selected track number
    '* Returns: the current selected/playing track
    
    'return the current track number
    lngReturnValues = mciSendString("status cd current track", strReturnString, _
                    Len(strReturnString), 0)
        'if no error occured then
            If lngReturnValues = 0 Then
                'save the current track number in "GetCurrentTrack"
                GetCurrentTrack = CInt(strReturnString)
            End If
End Function

Public Function GetTotalNumberOfTracks() As Integer
    '* Purpose: Gets the total number of tracks on CD
    '* Returns: the total number of tracks on CD
    
    'return the total number of track on the CD
    lngReturnValues = mciSendString("status cd number of tracks", strReturnString, _
                    Len(strReturnString), 0)
        'if no error then
            If lngReturnValues = 0 Then
                'save the total number of track on CD in "GetTotalNumberOfTracks"
                GetTotalNumberOfTracks = CInt(strReturnString)
            End If
End Function

Public Function GetCDPosition() As String
    '* Purpose: Gets the position of CD in HH:MM:SS:FF format
    '* Returns: the position of the CD
    
    'set the CD position in Track:Minutes:Seconds:Frames format
    SetCDFormat_TMSF
    'get the CD position
    lngReturnValues = mciSendString("status cd position wait", strReturnString, _
                    Len(strReturnString), 0)
        'if no error occured then
        If lngReturnValues = 0 Then
            'save the CD position in "GetCDPosition"
            GetCDPosition = strReturnString
        End If
End Function
