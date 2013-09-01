VERSION 5.00
Begin VB.Form 学号1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "抽学号 by 蚯蚓"
   ClientHeight    =   4884
   ClientLeft      =   120
   ClientTop       =   744
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   36
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "学号"
   MaxButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   3480
      Top             =   3000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "倒计时5秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4680
      TabIndex        =   2
      Top             =   3600
      Width           =   2892
   End
   Begin VB.CommandButton Command1 
      Caption         =   "控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Width           =   732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2772
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   3252
   End
   Begin VB.Menu xianshiname 
      Caption         =   "显示姓名"
   End
   Begin VB.Menu fiveperson 
      Caption         =   "5个人"
   End
   Begin VB.Menu about 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "学号1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Visible = False
Label1 = "1"
Timer1.Interval = 1
Timer2.Interval = 1000
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Font = "宋体"
Label1.FontSize = 150
Label1.ForeColor = vbRed
Command1.Caption = "Start"
Label2 = "5"
End Sub
Private Sub Command1_Click()

If Timer1.Enabled = True Then
Command1.Caption = "Start"
Timer1.Enabled = False
Else
Command1.Caption = "Stop"
Timer1.Enabled = True
End If

End Sub
Private Sub Command2_Click()
Label2.Visible = True
Timer1.Enabled = True

If Timer2.Enabled = True Then
Timer2.Enabled = False
Else
Timer2.Enabled = True
End If

If Timer2 = True Then
Command1.Enabled = False
Command2.Enabled = False
End If

End Sub
Private Sub Timer1_Timer()
Label1 = Label1 + 1

If Label1 = 65 Then
Label1 = "1"
End If

End Sub
Private Sub Timer2_Timer()
Label2 = Label2 - 1

If Label2 = 0 Then
Command1.Enabled = True
Command2.Enabled = True
Timer2.Enabled = False
Timer1.Enabled = False
Label2 = "5"
Label2.Visible = False
End If

End Sub
Private Sub xianshiname_Click()
学号1.Hide
学号2.Show
End Sub
Private Sub about_Click()
frmabout.Show
End Sub
Private Sub fiveperson_Click()
学号1.Hide
学号3.Show
End Sub
