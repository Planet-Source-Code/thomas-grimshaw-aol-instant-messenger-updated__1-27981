VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProfile 
   Caption         =   "Set Profile"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wait.."
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Please wait while I apply the changes..."
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Font -->"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Colour"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Underline"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Italics"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bold"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"frmProfile.frx":0000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please wait a tick..."
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   $"frmProfile.frx":007A
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu fonts 
      Caption         =   "&Fonts"
      Visible         =   0   'False
      Begin VB.Menu size 
         Caption         =   "&Size.."
         Begin VB.Menu f8pt 
            Caption         =   "8pt"
         End
         Begin VB.Menu f10pt 
            Caption         =   "10pt"
         End
         Begin VB.Menu f12pt 
            Caption         =   "12pt"
         End
         Begin VB.Menu f24pt 
            Caption         =   "24pt"
         End
         Begin VB.Menu f48pt 
            Caption         =   "48pt"
         End
      End
      Begin VB.Menu verdana 
         Caption         =   "&Verdana"
      End
      Begin VB.Menu Fixedsys 
         Caption         =   "&Fixedsys"
      End
      Begin VB.Menu garamond 
         Caption         =   "&Garamond"
      End
      Begin VB.Menu CourierNew 
         Caption         =   "&Courier New"
      End
      Begin VB.Menu ll 
         Caption         =   "-"
      End
      Begin VB.Menu addmorefonts 
         Caption         =   "&Add More Fonts"
      End
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addmorefonts_Click()
MsgBox "All the font items are in this form, under the menu Fonts. All you need to do is go into the menu editor, make a new item, and then copy the code from another font (Say CourierNew) and modify it to relflect your own font. <Easy really>."

End Sub

Private Sub Command1_Click()
Frame1.Visible = True
Command3.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Me.Refresh

l = RichToHTML(RichTextBox1, 0, Len(RichTextBox1.Text))
Call SendProc(2, "toc_set_info " & Chr(34) & Normalize(l) & Chr(34) & Chr(0))

Do Until Len(l) = 0
kj = kj + 1
p = bSetRegValue(HKEY_LOCAL_MACHINE, "Software\vbAIM Example", "Profile" + Format$(kj), Left$(l, 250))
If Len(l) > 250 Then l = Right$(l, Len(l) - 250) Else l = ""
Loop
l = bSetRegValue(HKEY_LOCAL_MACHINE, "Software\vbAIM Example", "Profile", Format$(kj))
Unload frmProfile

End Sub

Private Sub Command2_Click()
Unload frmProfile

End Sub

Private Sub Command3_Click()
RichTextBox1.SelBold = Not RichTextBox1.SelBold


End Sub

Private Sub Command4_Click()
RichTextBox1.SelItalic = Not RichTextBox1.SelItalic

End Sub

Private Sub Command5_Click()
RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline

End Sub

Private Sub Command6_Click()
CommonDialog1.Action = 3
RichTextBox1.SelColor = CommonDialog1.Color

End Sub

Private Sub Command7_Click()
PopupMenu fonts

End Sub

Private Sub CourierNew_Click()
RichTextBox1.SelFontName = "Courier New"

End Sub

Private Sub f10pt_Click()
RichTextBox1.SelFontSize = 10
End Sub

Private Sub f12pt_Click()
RichTextBox1.SelFontSize = 12
End Sub

Private Sub f24pt_Click()
RichTextBox1.SelFontSize = 24
End Sub

Private Sub f48pt_Click()
RichTextBox1.SelFontSize = 48
End Sub

Private Sub f8pt_Click()
RichTextBox1.SelFontSize = 8
End Sub

Private Sub Fixedsys_Click()
RichTextBox1.SelFontName = "fixedsys"
End Sub

Private Sub Form_Load()
On Error Resume Next

hjg = bGetRegValue("Software\vbAIM Example", "Profile")
If hjg = "" Then Me.Show: MsgBox "Registry storage of profile /  Update in HandleProc needs some work. Registry appears NOT to have profile stored. Unloading!": Timer1.Enabled = True: Exit Sub


iu$ = ""
        For X = 1 To Val(hjg)
        
        hjg2 = bGetRegValue("Software\vbAIM Example", "Profile" + Format$(X))
        hjg3 = ""
        For Y = 1 To Len(hjg2)
        If Mid$(hjg2, Y, 1) <> "\" Then hjg3 = hjg3 + Mid$(hjg2, Y, 1)
        Next Y
        iu$ = iu$ + hjg3
        Next X
Me.Show
Me.Refresh
Call HTMLToRich(iu$, RichTextBox1)
RichTextBox1.Visible = True

End Sub

Private Sub garamond_Click()
RichTextBox1.SelFontName = "Garamond"
End Sub

Private Sub Timer1_Timer()
Unload Me

End Sub

Private Sub verdana_Click()
RichTextBox1.SelFontName = "Verdana"
End Sub
