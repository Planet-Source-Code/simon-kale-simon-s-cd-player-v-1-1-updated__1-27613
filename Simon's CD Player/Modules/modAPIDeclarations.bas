Attribute VB_Name = "modAPIDeclarations"
Option Explicit

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
(ByVal nDrive As String) As Long

Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal _
nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" _
(ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

