VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      Height          =   2190
      Left            =   240
      Picture         =   "frmAbout.frx":1272
      ScaleHeight     =   2130
      ScaleWidth      =   2730
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2790
   End
   Begin VB.Label Label5 
      Caption         =   "WinAIM Passwords:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "ChiChis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1920
      TabIndex        =   18
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Not Yet Fu;lly Implemented"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   105
      TabIndex        =   17
      Top             =   3450
      Width           =   1530
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Tom Grimshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1920
      TabIndex        =   16
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Formatted Msg's: *"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Tom Grimshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1920
      TabIndex        =   14
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   """Profile"" feature:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Tom Grimshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   """Warning"" feature:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Tom Grimshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Steve"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   """Get Info"" feature:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Buddy list fix:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Protocol Fix/Update:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Original Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   """RTF2HTML"" and ""HTML2RTF"" code stolen with love from Joseph Huntley. Code fixed by Tom."
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   6975
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Chad J. Cox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Caption         =   "Tom Grimshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  lblInfo(0).Caption = "This project was developed by Chad J. Cox of www.dosfx.com. This is not a full client as some of the protocol  and a few features were left out. The reason being that this is meant to be an example only and is in no way what I would consider to be a full client ready for release." & vbCrLf & "Special thanks goes out to Pre (pre@dosfx.com). He has worked just as hard on this protocol and without a little teamwork, this project might not have come about." & vbCrLf & "If you have any questions or comments, please feel free to contact me (with the understanding that I do not have the time to teach you the protocol)." & vbCrLf & vbCrLf & "Also be sure to visit us in #visualbasic on irc.otherside.com."
  picLogo.Height = 1335
  
End Sub

