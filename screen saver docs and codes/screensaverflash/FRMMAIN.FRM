VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8970
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sf 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   12015
      _cx             =   4215497
      _cy             =   4202797
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6240
      Top             =   120
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   2040
      Top             =   8160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   1200
      Top             =   8160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   360
      Top             =   8160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1935
      Left            =   2760
      Shape           =   2  'Oval
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11775
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim boolvalue As Boolean
'########################
'author: somdutt ganguly
'address: gangulysomdutt@yahoo.com
'no 6, chandrodaya apt, bhaikaka nagar, thaltej, ahmedabad, gujarat, india - 380059
'Status: presently in II nd semister MCA - SV institute of computer studies, kadi, Gujarat
'North Gujarat University
'date: feb - 2003
'here make sure you keep the flash .swf file within the same folder
'this is a good screen saver using macromedia flash
'#########################

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 CursorOn
 End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 CursorOn
 End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 CursorOn
 End
End Sub


Private Sub Form_Load()
boolvalue = True
sf.Movie = App.Path & "\" & "dancefun.swf"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CursorOn
 End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If boolvalue = False Then
CursorOn
End
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CursorOn
 End
End Sub



Private Sub Form_Unload(Cancel As Integer)
 CursorOn
 On Error Resume Next
'Unload frmmain
'On Error Resume Next
Unload Me
On Error Resume Next
 End
End Sub



Private Sub Timer1_Timer()
boolvalue = False
If Shape1.Height = 1935 Then
Shape1.Height = 1000

Else
Shape1.Height = 1935
End If

On Error GoTo m
lbl1.Caption = lbl1.Caption + 1
Exit Sub
m:
lbl1.Caption = 0

End Sub




