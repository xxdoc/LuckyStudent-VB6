VERSION 5.00
Begin VB.Form ѧ��2 
   BorderStyle     =   0  'None
   Caption         =   "��ѧ�� by ���"
   ClientHeight    =   5448
   ClientLeft      =   120
   ClientTop       =   744
   ClientWidth     =   10788
   Icon            =   "ѧ��2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5448
   ScaleWidth      =   10788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer3 
      Left            =   480
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ʱ5��"
      Height          =   1092
      Left            =   6000
      TabIndex        =   3
      Top             =   4080
      Width           =   3252
   End
   Begin VB.Timer Timer 
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   1095
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label dao 
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
      Height          =   852
      Left            =   4920
      TabIndex        =   4
      Top             =   4320
      Width           =   612
   End
   Begin VB.Label numbersname 
      Height          =   3495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label number 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Menu xingming 
      Caption         =   "��������"
   End
   Begin VB.Menu oneperson 
      Caption         =   "1����"
   End
   Begin VB.Menu fiveperson 
      Caption         =   "5����"
   End
   Begin VB.Menu about 
      Caption         =   "����"
   End
End
Attribute VB_Name = "ѧ��2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dao.Visible = False
number = "6"
Label1 = "5"
Timer.Interval = 1
Timer2.Interval = 1000
Timer3.Interval = 1000
Timer2.Enabled = False
Timer3.Enabled = False
Timer.Enabled = True
number.Font = "����"
number.FontSize = 150
number.ForeColor = vbRed
numbersname.Font = "����"
numbersname.FontSize = 90
numbersname.ForeColor = vbRed
Command1.Caption = "Start"
End Sub
Private Sub Command2_Click()
dao.Visible = True
numbersname.Visible = False
Timer.Enabled = True
Timer3.Enabled = True
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

Private Sub Command1_Click()
Command1.Caption = "Start"
If Timer.Enabled = True Then
Timer.Enabled = False
numbersname.Visible = True
Else
Command1.Caption = "Stop"
Timer.Enabled = True
numbersname.Visible = False
End If


Select Case number
Case 1
numbersname.Caption = "����Ө"
Case 2
numbersname.Caption = "�����"
Case 3
numbersname.Caption = "�°ؾ�"
Case 4
numbersname.Caption = "�º��"
Case 5
numbersname.Caption = "��˼�"
Case 6
numbersname.Caption = "��־ҵ"
Case 7
numbersname.Caption = "��ϣ��"
Case 8
numbersname.Caption = "��Ӿ��"
Case 9
numbersname.Caption = "������"
Case 10
numbersname.Caption = "��С�"
Case 11
numbersname.Caption = "�γ�"
Case 12
numbersname.Caption = "�����"
Case 13
numbersname.Caption = "������"
Case 14
numbersname.Caption = "�ư���"
Case 15
numbersname.Caption = "�ƺ��"
Case 16
numbersname.Caption = "����ӱ"
Case 17
numbersname.Caption = "������"
Case 18
numbersname.Caption = "��¶��"
Case 19
numbersname.Caption = "������"
Case 20
numbersname.Caption = "������"
Case 21
numbersname.Caption = "��ʫ�"
Case 22
numbersname.Caption = "������"
Case 23
numbersname.Caption = "�γ��"
Case 24
numbersname.Caption = "����Ө"
Case 25
numbersname.Caption = "������"
Case 26
numbersname.Caption = "������"
Case 27
numbersname.Caption = "��ʫ��"
Case 28
numbersname.Caption = "������"
Case 29
numbersname.Caption = "����"
Case 30
numbersname.Caption = "�����"
Case 31
numbersname.Caption = "������"
Case 32
numbersname.Caption = "�����"
Case 33
numbersname.Caption = "Ī�ü�"
Case 34
numbersname.Caption = "Ī����"
Case 35
numbersname.Caption = "Ī�ӳ�"
Case 36
numbersname.Caption = "ŷ����"
Case 37
numbersname.Caption = "�����"
Case 38
numbersname.Caption = "��ΰ��"
Case 39
numbersname.Caption = "��ӱ��"
Case 40
numbersname.Caption = "������"
Case 41
numbersname.Caption = "�ۺ���"
Case 42
numbersname.Caption = "�մ���"
Case 43
numbersname.Caption = "̸С��"
Case 44
numbersname.Caption = "��о��"
Case 45
numbersname.Caption = "̷ӱ��"
Case 46
numbersname.Caption = "������"
Case 47
numbersname.Caption = "�¼���"
Case 48
numbersname.Caption = "��־�"
Case 49
numbersname.Caption = "�Ŀ���"
Case 50
numbersname.Caption = "������"
Case 51
numbersname.Caption = "�����"
Case 52
numbersname.Caption = "������"
Case 53
numbersname.Caption = "������"
Case 54
numbersname.Caption = "���㻪"
Case 55
numbersname.Caption = "�����"
Case 56
numbersname.Caption = "������"
Case 57
numbersname.Caption = "����"
Case 58
numbersname.Caption = "������"
Case 59
numbersname.Caption = "������"
Case 60
numbersname.Caption = "��ӱɺ"
Case 61
numbersname.Caption = "�Լ���"
Case 62
numbersname.Caption = "֣��Ȼ"
Case 63
numbersname.Caption = "�ӻ۾�"
Case 64
numbersname.Caption = "����ε"
End Select

End Sub
Private Sub Timer_Timer()
number = number + 1

If number = 65 Then
number = "1"
End If

End Sub
Private Sub Timer2_Timer()
dao = dao - 1

If dao = 0 Then
numbersname.Visible = True
Command1.Enabled = True
Command2.Enabled = True
Timer2.Enabled = False
Timer.Enabled = False
dao = "5"
dao.Visible = False
End If

End Sub

Private Sub Timer3_Timer()
Label1 = Label1 - 1
If Label1 = 0 Then
Timer3.Enabled = False
Label1 = "5"

Select Case number
Case 1
numbersname.Caption = "����Ө"
Case 2
numbersname.Caption = "�����"
Case 3
numbersname.Caption = "�°ؾ�"
Case 4
numbersname.Caption = "�º��"
Case 5
numbersname.Caption = "��˼�"
Case 6
numbersname.Caption = "��־ҵ"
Case 7
numbersname.Caption = "��ϣ��"
Case 8
numbersname.Caption = "��Ӿ��"
Case 9
numbersname.Caption = "������"
Case 10
numbersname.Caption = "��С�"
Case 11
numbersname.Caption = "�γ�"
Case 12
numbersname.Caption = "�����"
Case 13
numbersname.Caption = "������"
Case 14
numbersname.Caption = "�ư���"
Case 15
numbersname.Caption = "�ƺ��"
Case 16
numbersname.Caption = "����ӱ"
Case 17
numbersname.Caption = "������"
Case 18
numbersname.Caption = "��¶��"
Case 19
numbersname.Caption = "������"
Case 20
numbersname.Caption = "������"
Case 21
numbersname.Caption = "��ʫ�"
Case 22
numbersname.Caption = "������"
Case 23
numbersname.Caption = "�γ��"
Case 24
numbersname.Caption = "����Ө"
Case 25
numbersname.Caption = "������"
Case 26
numbersname.Caption = "������"
Case 27
numbersname.Caption = "��ʫ��"
Case 28
numbersname.Caption = "������"
Case 29
numbersname.Caption = "����"
Case 30
numbersname.Caption = "�����"
Case 31
numbersname.Caption = "������"
Case 32
numbersname.Caption = "�����"
Case 33
numbersname.Caption = "Ī�ü�"
Case 34
numbersname.Caption = "Ī����"
Case 35
numbersname.Caption = "Ī�ӳ�"
Case 36
numbersname.Caption = "ŷ����"
Case 37
numbersname.Caption = "�����"
Case 38
numbersname.Caption = "��ΰ��"
Case 39
numbersname.Caption = "��ӱ��"
Case 40
numbersname.Caption = "������"
Case 41
numbersname.Caption = "�ۺ���"
Case 42
numbersname.Caption = "�մ�־"
Case 43
numbersname.Caption = "̸С��"
Case 44
numbersname.Caption = "��о��"
Case 45
numbersname.Caption = "̷ӱ��"
Case 46
numbersname.Caption = "������"
Case 47
numbersname.Caption = "�¼���"
Case 48
numbersname.Caption = "��־�"
Case 49
numbersname.Caption = "�Ŀ���"
Case 50
numbersname.Caption = "������"
Case 51
numbersname.Caption = "�����"
Case 52
numbersname.Caption = "������"
Case 53
numbersname.Caption = "������"
Case 54
numbersname.Caption = "���㻪"
Case 55
numbersname.Caption = "�����"
Case 56
numbersname.Caption = "������"
Case 57
numbersname.Caption = "����"
Case 58
numbersname.Caption = "������"
Case 59
numbersname.Caption = "������"
Case 60
numbersname.Caption = "��ӱɺ"
Case 61
numbersname.Caption = "�Լ���"
Case 62
numbersname.Caption = "֣��Ȼ"
Case 63
numbersname.Caption = "�ӻ۾�"
Case 64
numbersname.Caption = "����ε"
End Select

End If
End Sub
Private Sub xingming_Click()
ѧ��1.Show
ѧ��2.Hide
End Sub
Private Sub about_Click()
frmabout.Show
End Sub
Private Sub oneperson_Click()
ѧ��1.Show
ѧ��2.Hide
End Sub
Private Sub fiveperson_Click()
ѧ��2.Hide
ѧ��3.Show
End Sub
