VERSION 5.00
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
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1680
      Top             =   1200
   End
   Begin VB.Timer timerman 
      Interval        =   1000
      Left            =   3480
      Top             =   1560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2055
      Left            =   3600
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label lbllabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I LOVE U ALL!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   5040
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim boolvalue As Boolean

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
Dim setting As String
setting = GetSetting(App.Title, "startup", "message", "hello")
If Len(Trim(setting)) = 0 Then
lbllabel.Caption = "I LOVE YOU!!"
Else
lbllabel.Caption = setting
End If
boolvalue = True

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
Shape2.Left = Rnd(Me.Height) * Me.Height
Shape2.Top = Rnd(Me.Height) * Me.Height

End Sub

Private Sub timerman_Timer()
If lbllabel.ForeColor = vbWhite Then
lbllabel.ForeColor = vbRed
Else
lbllabel.ForeColor = vbWhite
End If
boolvalue = False

End Sub


