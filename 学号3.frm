VERSION 5.00
Begin VB.Form ѧ��3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "��ѧ��  by  ���"
   ClientHeight    =   5064
   ClientLeft      =   120
   ClientTop       =   744
   ClientWidth     =   12552
   Icon            =   "ѧ��3.frx":0000
   LinkTopic       =   "ѧ��3"
   MaxButton       =   0   'False
   ScaleHeight     =   5064
   ScaleWidth      =   12552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer6 
      Left            =   5280
      Top             =   4080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ʱ5��"
      Height          =   1092
      Left            =   7560
      TabIndex        =   6
      Top             =   3600
      Width           =   2532
   End
   Begin VB.Timer Timer5 
      Left            =   10920
      Top             =   3240
   End
   Begin VB.Timer Timer4 
      Left            =   8640
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Left            =   6120
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Left            =   3480
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1092
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   2532
   End
   Begin VB.Label dao 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5640
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label number5 
      Height          =   2400
      Left            =   10080
      TabIndex        =   4
      Top             =   360
      Width           =   2100
   End
   Begin VB.Label number4 
      Height          =   2400
      Left            =   7680
      TabIndex        =   3
      Top             =   360
      Width           =   2100
   End
   Begin VB.Label number3 
      Height          =   2400
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label number2 
      Height          =   2400
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2100
   End
   Begin VB.Label number1 
      Height          =   2400
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2100
   End
   Begin VB.Menu xianshiname 
      Caption         =   "��ʾ����"
   End
   Begin VB.Menu oneperson 
      Caption         =   "1����"
   End
   Begin VB.Menu about 
      Caption         =   "����"
   End
End
Attribute VB_Name = "ѧ��3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
number1 = "1"
number2 = "22"
number3 = "35"
number4 = "54"
number5 = "52"
Timer1.Interval = 1
Timer1.Enabled = True
Timer2.Interval = 1
Timer2.Enabled = True
Timer3.Interval = 1
Timer3.Enabled = True
Timer4.Interval = 1
Timer4.Enabled = True
Timer5.Interval = 1
Timer5.Enabled = True
Timer6.Enabled = False
Timer6.Interval = 1000
number1.Font = "����"
number1.FontSize = 110
number1.ForeColor = vbRed
number2.Font = "����"
number2.FontSize = 110
number2.ForeColor = vbBlack
number3.Font = "����"
number3.FontSize = 110
number3.ForeColor = vbRed
number4.Font = "����"
number4.FontSize = 110
number4.ForeColor = vbBlack
number5.Font = "����"
number5.FontSize = 110
number5.ForeColor = vbRed
Command1.Caption = "Stop"
End Sub
Private Sub Command1_Click()

If Timer1.Enabled = True Then
Command1.Caption = "Start"
Timer1.Enabled = False
Else
Command1.Caption = "Stop"
Timer1.Enabled = True
End If

If Timer2.Enabled = True Then
Timer2.Enabled = False
Else
Timer2.Enabled = True
End If

If Timer3.Enabled = True Then
Timer3.Enabled = False
Else
Timer3.Enabled = True
End If

If Timer4.Enabled = True Then
Timer4.Enabled = False
Else
Timer4.Enabled = True
End If

If Timer5.Enabled = True Then
Timer5.Enabled = False
Else
Timer5.Enabled = True
End If
End Sub
Private Sub Command2_Click()
dao.Visible = True

Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True

If Timer6 = True Then
    Command1.Enabled = False
    Command2.Enabled = False
End If
End Sub
Private Sub Timer1_Timer()
number1 = number1 + 1

If number1 = 65 Then
number1 = "1"
End If

End Sub

Private Sub Timer2_Timer()
number2 = number2 + 1

If number2 = 65 Then
number2 = "1"
End If

End Sub

Private Sub Timer3_Timer()
number3 = number3 + 1

If number3 = 65 Then
number3 = "1"
End If

End Sub

Private Sub Timer4_Timer()
number4 = number4 + 1

If number4 = 65 Then
number4 = "1"
End If

End Sub

Private Sub Timer5_Timer()
number5 = number5 + 1

If number5 = 65 Then
number5 = "1"
End If

End Sub

Private Sub Timer6_Timer()
dao = dao - 1

If dao = 0 Then
Command1.Enabled = True
Command2.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
dao = "5"
dao.Visible = False
End If

End Sub
Private Sub about_Click()
frmabout.Show
End Sub
Private Sub oneperson_Click()
ѧ��1.Show
ѧ��3.Hide
End Sub
