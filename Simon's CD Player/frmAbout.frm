VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simon's CD Player"
   ClientHeight    =   3375
   ClientLeft      =   4410
   ClientTop       =   2595
   ClientWidth     =   5130
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2329.485
   ScaleMode       =   0  'User
   ScaleWidth      =   4817.335
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Unlimited (http://vbunlimited.cjb.net)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   840
      MouseIcon       =   "frmAbout.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Click here to visit my website"
      Top             =   1680
      Width           =   3450
   End
   Begin VB.Label label3 
      Caption         =   "Date: 28th September 2001"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblMail 
      Caption         =   "Email: simonkale@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   840
      MouseIcon       =   "frmAbout.frx":0620
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Click here to email me"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Written By: Simon Kale"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225.372
      X2              =   4507.448
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblDescription 
      Caption         =   "This is a MS Windows CD Player Clone."
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3525
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Caption         =   "Simon's CD Player"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225.372
      X2              =   4507.448
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.1"
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   3555
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":092A
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub lblMail_Click()
    ShellExecute Me.hwnd, "Open", "mailto:simonkale@yahoo.com", 0, 0, SW_SHOWNORMAL
End Sub

Private Sub lblWeb_Click()
    ShellExecute Me.hwnd, "Open", "http://vbunlimited.cjb.net/", 0, 0, SW_SHOWNORMAL
End Sub
