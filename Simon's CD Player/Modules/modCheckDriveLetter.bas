Attribute VB_Name = "modCheckDriveLetter"
Option Explicit
'This function is not written by me. I found it on the NET

Public Function FindCDDriveLetter() As String
'Use this function to discover which drive letter on a
'computer is the CDROM.
'Returns the drive letter in the form: "D:\" (without the
'quotes), or "None" if CDROM drive not present.
    Dim strDrives As String     'store list of drives
    Dim strThisDrive As String  'one drive from strDrives
    Dim intCounter As Integer
    Dim lngLenStr As Long       'length of returned string
    Dim lngDriveType As Long    'gets drive type as follows

    'Drive Types:
    '0: drive type cannot be determined
    '1: specified drive doesn't exist
    '2: removeable-disk can be removed, eg floppy or zip
    '3: fixed-disk cannot be removed, eg hard disk
    '4: remote-eg remote network drive
    '5: CD-ROM drive
    '6: RAM disk

    strDrives = Space$(255)

    lngLenStr = GetLogicalDriveStrings(255, strDrives)
    'strDrives has the names of all the root directories.
    'Each entry takes four characters-three for the name plus
    'a null character. String ends with a second null.
    'Example: A:\[Null]C:\[Null]D:\[Null][Null]

    For intCounter = 1 To lngLenStr Step 4
    'Count by fours to get the letter of each drive
        strThisDrive = Mid$(strDrives, intCounter, 3)
        lngDriveType = GetDriveType(strThisDrive)
        If lngDriveType = 5 Then 'It's a CDROM
            FindCDDriveLetter = UCase$(strThisDrive)
            Exit Function
        End If
    Next intCounter

    FindCDDriveLetter = "None" 'System doesn't have a CDROM
End Function


