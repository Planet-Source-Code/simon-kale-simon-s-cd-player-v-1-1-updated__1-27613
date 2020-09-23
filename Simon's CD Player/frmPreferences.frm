VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   2490
   ClientLeft      =   4590
   ClientTop       =   2790
   ClientWidth     =   4515
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton optLargeFont 
      Caption         =   "&Large font"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optSmallFont 
      Caption         =   "S&mall font"
      Height          =   435
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtIntroPlay 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "10"
      Top             =   960
      Width           =   350
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display font"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4215
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "[01] 00:00"
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
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.VScrollBar vsbCDPlayer 
      Height          =   255
      Left            =   450
      Max             =   5
      Min             =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Value           =   5
      Width           =   250
   End
   Begin VB.CheckBox chkShowToolTips 
      Caption         =   "Show &tool tips"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.CheckBox chkSaveSettings 
      Caption         =   "&Save settings on exit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.CheckBox chkStopCDPlay 
      Caption         =   "Stop &CD playing on exit"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblIntoPlay 
      Caption         =   "&Intro play length(seconds)"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   975
      Width           =   2055
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    'get outta here
    Unload Me
End Sub

Public Sub Form_Load()
    'the initial value of "Intro Play Length" as 10 seconds
    txtIntroPlay.Text = 10
    'get the states of all check boxes (i.e. if they are checked or not)
    Call GetFromRegistry
End Sub

Private Sub OKButton_Click()
    'check if the user has entered Intro Play Length less than 4 and greater than 60
    If Int(txtIntroPlay.Text) < 4 Or Int(txtIntroPlay.Text) > 60 Then
        'then show a message box telling the user he can't do that
        MsgBox "Intro Play Length must be greater than 4 and less than 59 seconds.", vbOKOnly + vbApplicationModal, "Simon's CD Player"
        'set the default value of 15 as Intro Play Length
        txtIntroPlay.Text = "15"
    End If
    'save changes to registry
    Call SaveToRegistry
    'get outta here
    Unload Me
End Sub

Private Sub optLargeFont_Click()
    'change the "txtTime" (in frmPreferences; it is "lblTime") display font to large
    lblTime.FontSize = 18
    frmCdplay.txtTime.FontSize = 18
End Sub

Private Sub optSmallFont_Click()
    'change the "txtTime" (in frmPreferences; it is "lblTime") display font to small
    lblTime.FontSize = 15
    frmCdplay.txtTime.FontSize = 15
End Sub

Private Sub txtIntroPlay_KeyPress(KeyAscii As Integer)
    'don't allow the user to type letters in the text box
    If KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Then
        Exit Sub
    End If
    
    'allow only 2 numbers to be entered as "Intro Play Length"
    SendMessage txtIntroPlay.hwnd, EM_LIMITTEXT, 2, ByVal 0

End Sub

Private Sub vsbCDPlayer_Change()
    'set "Intro Play Length" value equal to vertical scroll bar value
    txtIntroPlay.Text = Str(vsbCDPlayer.Value)
End Sub
