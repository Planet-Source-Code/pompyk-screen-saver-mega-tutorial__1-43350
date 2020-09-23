VERSION 5.00
Begin VB.Form frmsettings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Screensaver Settings"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framesetting 
      BackColor       =   &H00C0C0C0&
      Caption         =   "S E T T I N G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdclose 
         Caption         =   "SAVE SETTING"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "I LOVE YOU!!"
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdclose_Click()
SaveSetting App.Title, "startup", "message", Text1.Text

'On Error Resume Next
'Unload frmmain
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'Unload frmmain
On Error Resume Next
Unload Me
End Sub

